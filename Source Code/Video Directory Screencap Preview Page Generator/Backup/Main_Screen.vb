'Imports Microsoft.Win32
Imports System.IO



Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False
    Private application_exit As Boolean = False
    Private shutting_down As Boolean = False
    Public WorkingFolder As String = ""
    Private CurrentListItem As Integer = 0


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
       
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
       
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
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents example2 As System.Windows.Forms.PictureBox
    Friend WithEvents example3 As System.Windows.Forms.PictureBox
    Friend WithEvents example6 As System.Windows.Forms.PictureBox
    Friend WithEvents example4 As System.Windows.Forms.PictureBox
    Friend WithEvents example7 As System.Windows.Forms.PictureBox
    Friend WithEvents example5 As System.Windows.Forms.PictureBox
    Friend WithEvents example8 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents example1 As System.Windows.Forms.Label
    Friend WithEvents example9 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents VideoFilesFound As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents example10 As System.Windows.Forms.Label
    Friend WithEvents MediaPlayer1 As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CurrentPositionTimer As System.Windows.Forms.Timer
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents ContextMenu3 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents Label14 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Label8 = New System.Windows.Forms.Label
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.example2 = New System.Windows.Forms.PictureBox
        Me.example3 = New System.Windows.Forms.PictureBox
        Me.example6 = New System.Windows.Forms.PictureBox
        Me.example4 = New System.Windows.Forms.PictureBox
        Me.example7 = New System.Windows.Forms.PictureBox
        Me.example5 = New System.Windows.Forms.PictureBox
        Me.example8 = New System.Windows.Forms.PictureBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.example10 = New System.Windows.Forms.Label
        Me.example1 = New System.Windows.Forms.Label
        Me.example9 = New System.Windows.Forms.Label
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.VideoFilesFound = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.MediaPlayer1 = New AxWMPLib.AxWindowsMediaPlayer
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CurrentPositionTimer = New System.Windows.Forms.Timer(Me.components)
        Me.ContextMenu3 = New System.Windows.Forms.ContextMenu
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.Label14 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.MediaPlayer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label8.Location = New System.Drawing.Point(904, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 16)
        Me.Label8.TabIndex = 33
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label8, "Current System Time")
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Resting..."
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem5, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Main Screen"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "Refresh Main Screen"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 3
        Me.MenuItem3.Text = "Exit Application"
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Silver
        Me.Label9.Location = New System.Drawing.Point(904, 654)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(112, 16)
        Me.Label9.TabIndex = 65
        Me.Label9.Text = "BUILD 20060223.3"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label9, "Application Build Number")
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 224)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 16)
        Me.Label3.TabIndex = 72
        Me.Label3.Text = "Play"
        Me.ToolTip1.SetToolTip(Me.Label3, "Play")
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(48, 224)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 16)
        Me.Label4.TabIndex = 73
        Me.Label4.Text = "Stop"
        Me.ToolTip1.SetToolTip(Me.Label4, "Stop")
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(80, 224)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(44, 16)
        Me.Label5.TabIndex = 74
        Me.Label5.Text = "Pause"
        Me.ToolTip1.SetToolTip(Me.Label5, "Pause")
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(176, 224)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(14, 16)
        Me.Label6.TabIndex = 75
        Me.Label6.Text = ">"
        Me.ToolTip1.SetToolTip(Me.Label6, "Skip Forward Small")
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(152, 224)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(16, 16)
        Me.Label7.TabIndex = 76
        Me.Label7.Text = "<"
        Me.ToolTip1.SetToolTip(Me.Label7, "Skip Backwards Small")
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(24, 208)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(184, 16)
        Me.Label10.TabIndex = 77
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.Label10, "Current Position")
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(128, 224)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(18, 16)
        Me.Label11.TabIndex = 78
        Me.Label11.Text = "<<"
        Me.ToolTip1.SetToolTip(Me.Label11, "Skip Backwords Large")
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(192, 224)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(18, 16)
        Me.Label12.TabIndex = 79
        Me.Label12.Text = ">>"
        Me.ToolTip1.SetToolTip(Me.Label12, "Skip Forward Large")
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.DarkGreen
        Me.Label13.Location = New System.Drawing.Point(16, 416)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(200, 64)
        Me.Label13.TabIndex = 81
        Me.ToolTip1.SetToolTip(Me.Label13, "Current Position")
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(648, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(144, 72)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 55
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(608, 40)
        Me.Label1.TabIndex = 56
        Me.Label1.Text = "Video Menu"
        '
        'example2
        '
        Me.example2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.example2.Enabled = False
        Me.example2.Location = New System.Drawing.Point(24, 48)
        Me.example2.Name = "example2"
        Me.example2.Size = New System.Drawing.Size(288, 224)
        Me.example2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.example2.TabIndex = 57
        Me.example2.TabStop = False
        Me.example2.Tag = "example"
        Me.example2.Visible = False
        '
        'example3
        '
        Me.example3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.example3.Enabled = False
        Me.example3.Location = New System.Drawing.Point(328, 48)
        Me.example3.Name = "example3"
        Me.example3.Size = New System.Drawing.Size(128, 96)
        Me.example3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.example3.TabIndex = 58
        Me.example3.TabStop = False
        Me.example3.Tag = "example"
        Me.example3.Visible = False
        '
        'example6
        '
        Me.example6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.example6.Enabled = False
        Me.example6.Location = New System.Drawing.Point(328, 152)
        Me.example6.Name = "example6"
        Me.example6.Size = New System.Drawing.Size(128, 96)
        Me.example6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.example6.TabIndex = 59
        Me.example6.TabStop = False
        Me.example6.Tag = "example"
        Me.example6.Visible = False
        '
        'example4
        '
        Me.example4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.example4.Enabled = False
        Me.example4.Location = New System.Drawing.Point(472, 48)
        Me.example4.Name = "example4"
        Me.example4.Size = New System.Drawing.Size(128, 96)
        Me.example4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.example4.TabIndex = 60
        Me.example4.TabStop = False
        Me.example4.Tag = "example"
        Me.example4.Visible = False
        '
        'example7
        '
        Me.example7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.example7.Enabled = False
        Me.example7.Location = New System.Drawing.Point(472, 152)
        Me.example7.Name = "example7"
        Me.example7.Size = New System.Drawing.Size(128, 96)
        Me.example7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.example7.TabIndex = 61
        Me.example7.TabStop = False
        Me.example7.Tag = "example"
        Me.example7.Visible = False
        '
        'example5
        '
        Me.example5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.example5.Enabled = False
        Me.example5.Location = New System.Drawing.Point(616, 48)
        Me.example5.Name = "example5"
        Me.example5.Size = New System.Drawing.Size(128, 96)
        Me.example5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.example5.TabIndex = 62
        Me.example5.TabStop = False
        Me.example5.Tag = "example"
        Me.example5.Visible = False
        '
        'example8
        '
        Me.example8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.example8.Enabled = False
        Me.example8.Location = New System.Drawing.Point(616, 152)
        Me.example8.Name = "example8"
        Me.example8.Size = New System.Drawing.Size(128, 96)
        Me.example8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.example8.TabIndex = 63
        Me.example8.TabStop = False
        Me.example8.Tag = "example"
        Me.example8.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.Controls.Add(Me.example10)
        Me.Panel1.Controls.Add(Me.example1)
        Me.Panel1.Controls.Add(Me.example2)
        Me.Panel1.Controls.Add(Me.example4)
        Me.Panel1.Controls.Add(Me.example6)
        Me.Panel1.Controls.Add(Me.example8)
        Me.Panel1.Controls.Add(Me.example3)
        Me.Panel1.Controls.Add(Me.example7)
        Me.Panel1.Controls.Add(Me.example5)
        Me.Panel1.Controls.Add(Me.example9)
        Me.Panel1.Location = New System.Drawing.Point(0, 80)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(784, 564)
        Me.Panel1.TabIndex = 64
        '
        'example10
        '
        Me.example10.Enabled = False
        Me.example10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.example10.ForeColor = System.Drawing.Color.DarkSeaGreen
        Me.example10.Location = New System.Drawing.Point(568, 256)
        Me.example10.Name = "example10"
        Me.example10.Size = New System.Drawing.Size(184, 16)
        Me.example10.TabIndex = 66
        Me.example10.Tag = "example"
        Me.example10.Text = "(Right-click on Image to Play Video)"
        Me.example10.Visible = False
        '
        'example1
        '
        Me.example1.Enabled = False
        Me.example1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.example1.Location = New System.Drawing.Point(24, 16)
        Me.example1.Name = "example1"
        Me.example1.Size = New System.Drawing.Size(720, 23)
        Me.example1.TabIndex = 64
        Me.example1.Tag = "example"
        Me.example1.Text = "Title"
        Me.example1.Visible = False
        '
        'example9
        '
        Me.example9.Enabled = False
        Me.example9.Location = New System.Drawing.Point(328, 256)
        Me.example9.Name = "example9"
        Me.example9.Size = New System.Drawing.Size(240, 16)
        Me.example9.TabIndex = 65
        Me.example9.Tag = "example"
        Me.example9.Text = "File Details"
        Me.example9.Visible = False
        '
        'ContextMenu2
        '
        Me.ContextMenu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem4})
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Text = "Preview"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.Text = "Play"
        '
        'VideoFilesFound
        '
        Me.VideoFilesFound.BackColor = System.Drawing.Color.Transparent
        Me.VideoFilesFound.Location = New System.Drawing.Point(32, 56)
        Me.VideoFilesFound.Name = "VideoFilesFound"
        Me.VideoFilesFound.Size = New System.Drawing.Size(600, 23)
        Me.VideoFilesFound.TabIndex = 66
        Me.VideoFilesFound.Text = "0"
        '
        'Panel2
        '
        Me.Panel2.BackgroundImage = CType(resources.GetObject("Panel2.BackgroundImage"), System.Drawing.Image)
        Me.Panel2.Controls.Add(Me.VideoFilesFound)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(640, 80)
        Me.Panel2.TabIndex = 67
        '
        'MediaPlayer1
        '
        Me.MediaPlayer1.ContainingControl = Me
        Me.MediaPlayer1.Enabled = True
        Me.MediaPlayer1.Location = New System.Drawing.Point(16, 32)
        Me.MediaPlayer1.Name = "MediaPlayer1"
        Me.MediaPlayer1.OcxState = CType(resources.GetObject("MediaPlayer1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.MediaPlayer1.Size = New System.Drawing.Size(192, 176)
        Me.MediaPlayer1.TabIndex = 68
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BackgroundImage = CType(resources.GetObject("Panel3.BackgroundImage"), System.Drawing.Image)
        Me.Panel3.Controls.Add(Me.ListBox1)
        Me.Panel3.Controls.Add(Me.Label13)
        Me.Panel3.Controls.Add(Me.PictureBox2)
        Me.Panel3.Controls.Add(Me.Label12)
        Me.Panel3.Controls.Add(Me.Label11)
        Me.Panel3.Controls.Add(Me.Label10)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.MediaPlayer1)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.Label14)
        Me.Panel3.Location = New System.Drawing.Point(784, 56)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(232, 584)
        Me.Panel3.TabIndex = 70
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.BackColor = System.Drawing.Color.White
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListBox1.ContextMenu = Me.ContextMenu3
        Me.ListBox1.ForeColor = System.Drawing.Color.DarkGreen
        Me.ListBox1.Location = New System.Drawing.Point(16, 464)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.ScrollAlwaysVisible = True
        Me.ListBox1.Size = New System.Drawing.Size(192, 119)
        Me.ListBox1.TabIndex = 82
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Location = New System.Drawing.Point(16, 248)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(192, 160)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 80
        Me.PictureBox2.TabStop = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 16)
        Me.Label2.TabIndex = 71
        Me.Label2.Text = "Video Preview"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CurrentPositionTimer
        '
        Me.CurrentPositionTimer.Interval = 500
        '
        'ContextMenu3
        '
        Me.ContextMenu3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem7, Me.MenuItem8})
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 0
        Me.MenuItem7.Text = "Preview"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 1
        Me.MenuItem8.Text = "Play"
        '
        'Label14
        '
        Me.Label14.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label14.Location = New System.Drawing.Point(0, 544)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(240, 48)
        Me.Label14.TabIndex = 83
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(1018, 668)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel2)
        Me.ForeColor = System.Drawing.Color.DarkGreen
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Main_Screen"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Video Directory Screencap Preview Page Generator 1.0"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.MediaPlayer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Try

                
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If

                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                    filewriter.Close()
                Catch exp As Exception
                    Dim Display_Message2 As New Display_Message("Unable to create the disk-based error log")
                    Display_Message2.ShowDialog()
                End Try
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Video Directory Screencap Preview Page Generator's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            '
            'Label9.Text = MediaPlayer1.openState.ToString
            '
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
            If application_exit = True Then
                dataloaded = True
                splash_loader.Visible = False
                Me.Close()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Function FormatSize(ByVal SizeInBytes As Double) As String
        Try
            If SizeInBytes < 1024 Then
                Return String.Format("{0:N0} B", SizeInBytes)
            ElseIf SizeInBytes < 1024 * 1024 Then
                Return String.Format("{0:N2} KB", SizeInBytes / 1024)
            Else
                Return String.Format("{0:N2} MB", SizeInBytes / (1024 * 1024))
            End If
        Catch ex As Exception
            Error_Handler(ex)
            Return "Error"
        End Try
    End Function

    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False


            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter


            myProcess.StartInfo.CreateNoWindow = True

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            'myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()

        Catch ex As Exception
            Error_Handler(ex, "DosShellCommand")
        End Try

        Return s
    End Function

    Private Function ApplicationLauncher(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process
            myProcess.StartInfo.UseShellExecute = False
            Dim sErr As StreamReader
            Dim sOut As StreamReader
            Dim sIn As StreamWriter
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True

            myProcess.StartInfo.FileName = AppToRun

            myProcess.Start()
            sIn = myProcess.StandardInput
            sIn.AutoFlush = True

            sOut = myProcess.StandardOutput()
            sErr = myProcess.StandardError

            sIn.Write(AppToRun & System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()

            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()

        Catch ex As Exception
            Error_Handler(ex, "ApplicationLauncher")
        End Try
        Return s
    End Function

    Private Sub shuffle(ByRef arrayToBeShuffled As Array, ByVal numberOfTimesToShuffle As Integer)
        Try
            Dim rndPosition As New Random(DateTime.Now.Millisecond)

        For i As Integer = 1 To numberOfTimesToShuffle

            For i2 As Integer = 1 To arrayToBeShuffled.Length

                swap(arrayToBeShuffled(rndPosition.Next(0, arrayToBeShuffled.Length)), arrayToBeShuffled(rndPosition.Next(0, arrayToBeShuffled.Length)))

            Next i2

        Next i
        Catch ex As Exception
            Error_Handler(ex, "shuffle")
        End Try
    End Sub

    Private Sub swap(ByRef arg1 As Object, ByRef arg2 As Object)
        Try
            Dim strTemp As String

            strTemp = arg1

            arg1 = arg2

        arg2 = strTemp
        Catch ex As Exception
            Error_Handler(ex, "swap")
        End Try
    End Sub



    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Not WorkingFolder = "" Then
                Dim testinfo As DirectoryInfo = New DirectoryInfo(WorkingFolder)
                If testinfo.Exists = False Then
                    WorkingFolder = Application.StartupPath
                End If
                testinfo = Nothing
            Else
                WorkingFolder = Application.StartupPath
            End If

            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")

            Timer2.Start()


            application_exit = False

            Dim dinfo As DirectoryInfo = New DirectoryInfo(WorkingFolder)

            Dim tinfo As FileInfo = New FileInfo((dinfo.FullName & "\config.ini").Replace("\\", "\"))
            If tinfo.Exists = True Then
                Dim filereader As StreamReader = New StreamReader((dinfo.FullName & "\config.ini").Replace("\\", "\"))
                Dim lineread As String
                While filereader.Peek > -1
                    lineread = filereader.ReadLine().Trim
                    If lineread.StartsWith("Page_Title=") = True Then
                        Label1.Text = lineread.Replace("Page_Title=", "")
                        Me.Text = lineread.Replace("Page_Title=", "")
                        NotifyIcon1.Text = lineread.Replace("Page_Title=", "")
                    End If
                End While
                filereader.Close()
                filereader = Nothing
            End If
            tinfo = Nothing


            Dim dothedeed As Boolean = True
            Dim tested As DirectoryInfo
            For Each tested In dinfo.GetDirectories
                If tested.Name.ToLower = "thumbnails" Then
                    dothedeed = False
                    Exit For
                End If
            Next
            If dothedeed = True Then
                tested = dinfo.CreateSubdirectory("Thumbnails")
            End If


            Dim top1 As Long = example1.Top
            Dim top2 As Long = example2.Top
            Dim top3 As Long = example3.Top
            Dim top4 As Long = example4.Top
            Dim top5 As Long = example5.Top
            Dim top6 As Long = example6.Top
            Dim top7 As Long = example7.Top
            Dim top8 As Long = example8.Top
            Dim top9 As Long = example9.Top
            Dim top10 As Long = example10.Top

            Dim picorder As Integer() = {2, 3, 4, 5, 6, 7, 8}
            shuffle(picorder, 7)
            Dim i1 As Integer = 0
            Dim finfo As FileInfo

            Dim firstvideo As String = ""

            For Each finfo In dinfo.GetFiles()
                If finfo.Exists = True Then


                If finfo.Extension.ToLower = ".wmv" Or finfo.Extension.ToLower = ".avi" Or finfo.Extension.ToLower = ".mpg" Then
                        VideoFilesFound.Text = (CInt(VideoFilesFound.Text) + 1).ToString
                        ListBox1.Items.Add(finfo.Name.Replace(finfo.Extension, ""))
                        If firstvideo = "" Then
                            firstvideo = finfo.Name.Replace(finfo.Extension, "")
                        End If
                        Dim tops As Long() = {top1, top9, top10}
                        Dim lbls As Label() = {example1, example9, example10}
                        For i1 = 0 To 2
                            Dim lb1 As Label = New Label
                            lb1.Tag = finfo.FullName
                            lb1.ContextMenu = ContextMenu2
                            lb1.Top = tops(i1)
                            lb1.Height = lbls(i1).Height
                            lb1.Width = lbls(i1).Width
                            lb1.Left = lbls(i1).Left
                            lb1.Font = lbls(i1).Font
                            lb1.ForeColor = lbls(i1).ForeColor
                            If i1 = 0 Then
                                lb1.Text = finfo.Name.Replace(finfo.Extension, "")
                            End If
                            If i1 = 1 Then
                                lb1.Text = "Duration: " & "Unknown"
                                lb1.Text = lb1.Text & "; Size: " & FormatSize(finfo.Length)
                            End If
                            If i1 = 2 Then
                                lb1.Text = lbls(i1).Text
                            End If

                            lb1.Visible = True
                            Panel1.Controls.Add(lb1)
                            ToolTip1.SetToolTip(lb1, finfo.Name)
                        Next

                        Dim fil As FileInfo

                        For thumbrun As Integer = 1 To 9
                            fil = New FileInfo((tested.FullName & "\" & finfo.Name.Replace(finfo.Extension, "") & thumbrun & ".jpg").Replace("\\", "\"))
                            If fil.Exists = False Then
                                fil = Nothing
                                For thumbrun2 As Integer = 1 To 9
                                    fil = New FileInfo((tested.FullName & "\" & finfo.Name.Replace(finfo.Extension, "") & thumbrun2 & ".jpg").Replace("\\", "\"))
                                    If fil.Exists = True Then
                                        fil.Delete()
                                    End If
                                    fil = Nothing
                                Next
                                Dim apptorun As String
                                apptorun = ("""" & Application.StartupPath & "\SnatchIt.exe"" –thumbs 10 """ & finfo.FullName & """ """ & WorkingFolder & "\Thumbnails""").Replace("\\", "\")
                                ApplicationLauncher(apptorun)
                            End If
                            fil = Nothing
                        Next


                        fil = New FileInfo((tested.FullName & "\" & finfo.Name.Replace(finfo.Extension, "") & "1.jpg").Replace("\\", "\"))

                        Dim topp As Long() = {top2, top3, top4, top5, top6, top7, top8}
                        Dim lblp As PictureBox() = {example2, example3, example4, example5, example6, example7, example8}
                        For i1 = 0 To 6
                            Dim lb1 As PictureBox = New PictureBox
                            lb1.Tag = finfo.FullName
                            lb1.ContextMenu = ContextMenu2
                            lb1.Top = topp(i1)
                            lb1.Height = lblp(i1).Height
                            lb1.Width = lblp(i1).Width
                            lb1.Left = lblp(i1).Left
                            lb1.Font = lblp(i1).Font
                            lb1.Visible = True
                            lb1.BorderStyle = lblp(i1).BorderStyle
                            lb1.SizeMode = lblp(i1).SizeMode
                            lb1.Image = lb1.Image.FromFile((tested.FullName & "\" & finfo.Name.Replace(finfo.Extension, "") & picorder(i1) & ".jpg").Replace("\\", "\"))
                            Panel1.Controls.Add(lb1)
                            ToolTip1.SetToolTip(lb1, finfo.Name)
                        Next

                        'fil = New FileInfo((tested.FullName & "\" & finfo.Name.Replace(finfo.Extension, "") & "_Link.htm").Replace("\\", "\"))
                        'If fil.Exists = False Then
                        '    Dim filewriter As StreamWriter = New StreamWriter((tested.FullName & "\" & finfo.Name.Replace(finfo.Extension, "") & "_Link.htm").Replace("\\", "\"), False)
                        '    filewriter.WriteLine("<HTML><HEAD><TITLE>" & finfo.Name.Replace(finfo.Extension, "") & "</TITLE></HEAD>")
                        '    filewriter.WriteLine("<BODY><A HREF=""" & finfo.FullName & """ target=""_blank""><B>Play Movie</B></A></BODY></HTML>")
                        '    filewriter.Close()
                        '    filewriter = Nothing
                        'End If

                        fil = Nothing

                        top1 = top1 + 280
                        top2 = top2 + 280
                        top3 = top3 + 280
                        top4 = top4 + 280
                        top5 = top5 + 280
                        top6 = top6 + 280
                        top7 = top7 + 280
                        top8 = top8 + 280
                        top9 = top9 + 280
                        top10 = top10 + 280
                    End If
                End If
            Next
            VideoFilesFound.Text = VideoFilesFound.Text & " Video Files Located"
            If Not firstvideo = "" Then
                Load_Details(firstvideo)
            End If


            tested = Nothing
            finfo = Nothing
            dinfo = Nothing
            Me.WindowState = FormWindowState.Normal

            CurrentPositionTimer.Start()

            dataloaded = True
            splash_loader.Visible = False
            splash_loader.Opacity = 0
            splash_loader.WindowState = FormWindowState.Minimized
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub exit_application()
        Try
            Me.Opacity = 0
            Me.WindowState = FormWindowState.Minimized
            shutting_down = True
            If Not MediaPlayer1.currentMedia Is Nothing Then
                MediaPlayer1.Ctlcontrols.stop()
            End If
            CurrentPositionTimer.Stop()
            Timer2.Stop()
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex, "Shutting Down Application")
        Finally
            Application.Exit()
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub Main_Screen_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click


        Dim x, y As Integer
    
        x = Cursor.Position.X
        y = Cursor.Position.Y
    
        For Each ctrl As Control In Panel1.Controls
            Dim beginx As Integer = 0
            Dim endx As Integer = 0
            Dim beginy As Integer = 0
            Dim endy As Integer = 0
            beginx = ctrl.Left + Panel1.Left + Me.Left
            beginy = ctrl.Top + Panel1.Top + Me.Top
            endx = ctrl.Left + ctrl.Width + Panel1.Left + Me.Left
            endy = ctrl.Top + ctrl.Height + Panel1.Top + Me.Top

            If x >= beginx And x <= endx Then

                If y >= beginy And y <= endy Then
                    If (Not ctrl.Tag = "example") And (Not ctrl.Tag = "") Then
                        Dim fin As FileInfo = New FileInfo(ctrl.Tag)
                        Label13.Text = fin.Name
                        PictureBox2.Image = PictureBox2.Image.FromFile(((WorkingFolder & "\Thumbnails\" & fin.Name.Replace(fin.Extension, "") & "3.jpg").Replace("\\", "\")))
                        'MsgBox((Application.StartupPath & "\Thumbnails\" & fin.Name.Replace(fin.Extension, "") & "1.jpg").Replace("\\", "\"))
                        fin = Nothing
                        MediaPlayer1.URL = ctrl.Tag
                        MediaPlayer1.Ctlcontrols.stop()
                        DosShellCommand("""" & ctrl.Tag & """")

                        Exit For
                    End If

                End If
            End If


        Next
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Try
            Me.WindowState = FormWindowState.Minimized
            System.Diagnostics.Process.Start("""" & Application.ExecutablePath & """")
            Me.Close()
        Catch ex As Exception
            Error_Handler(ex, "Display Refresh")
        End Try
    End Sub

   

    'Private Sub Duration_Timer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Duration_Timer.Stop()
    '    For Each ctrl As Control In Panel1.Controls
    '        Label9.Text = MediaPlayer1.openState.ToString
    '        If ctrl.GetType.ToString = "System.Windows.Forms.Label" Then
    '            If ctrl.Text.StartsWith("Duration: Unknown") = True Then
    '                ''Dim vid As WMPLib.IWMPMedia = MediaPlayer1.newMedia(finfo.FullName)
    '                ''vid.setItemInfo()

    '                MediaPlayer1.URL = ctrl.Tag
    '                ''Dim timeout As Long = 1000000

    '                'While Not MediaPlayer1.openState.ToString = "wmposMediaOpen"

    '                'End While
    '                'MsgBox(ctrl.Tag)
    '                Label9.Text = MediaPlayer1.openState.ToString
    '                '    '   timeout = timeout - 1
    '                '    '  If timeout <= 0 Then
    '                '    ' Exit While
    '                '    'End If
    '                'End While
    '                'MediaPlayer1.Ctlcontrols.stop()
    '                If MediaPlayer1.openState.ToString = "wmposMediaOpen" Then
    '                    ctrl.Text = ctrl.Text.Replace("Unknown", MediaPlayer1.currentMedia.durationString)
    '                End If

    '                '  MsgBox(MediaPlayer1.currentMedia.durationString)

    '                ''End While
    '                ''Label9.Text = MediaPlayer1.playState.ToString
    '                ''End While
    '                ''Dim vid As WMPLib.IWMPMedia = MediaPlayer1.currentMedia

    '            End If

    '        End If


    '    Next


    '    Duration_Timer.Start()
    'End Sub

    Private Sub MediaPlayer1_OpenStateChange(ByVal sender As Object, ByVal e As AxWMPLib._WMPOCXEvents_OpenStateChangeEvent) Handles MediaPlayer1.OpenStateChange
        ' MsgBox(MediaPlayer1.openState.ToString)
        If MediaPlayer1.openState.ToString = "wmposMediaOpen" Then


        For Each ctrl As Control In Panel1.Controls
            ' Label9.Text = MediaPlayer1.openState.ToString
                If ctrl.GetType.ToString = "System.Windows.Forms.Label" Then
                    ' MsgBox(ctrl.Tag.ToString & " = " & MediaPlayer1.URL.ToString)
                    If ctrl.Tag = MediaPlayer1.URL Then
                        If ctrl.Text.StartsWith("Duration: Unknown") = True Then
                            ''Dim vid As WMPLib.IWMPMedia = MediaPlayer1.newMedia(finfo.FullName)
                            ''vid.setItemInfo()

                            ' MediaPlayer1.URL = ctrl.Tag
                            ''Dim timeout As Long = 1000000

                            'While Not MediaPlayer1.openState.ToString = "wmposMediaOpen"

                            'End While
                            'MsgBox(ctrl.Tag)
                            ' Label9.Text = MediaPlayer1.openState.ToString
                            '    '   timeout = timeout - 1
                            '    '  If timeout <= 0 Then
                            '    ' Exit While
                            '    'End If
                            'End While
                            'MediaPlayer1.Ctlcontrols.stop()
                            'If MediaPlayer1.openState.ToString = "wmposMediaOpen" Then
                            ctrl.Text = ctrl.Text.Replace("Unknown", MediaPlayer1.currentMedia.durationString)
                            ' End If

                            '  MsgBox(MediaPlayer1.currentMedia.durationString)

                            ''End While
                            ''Label9.Text = MediaPlayer1.playState.ToString
                            ''End While
                            ''Dim vid As WMPLib.IWMPMedia = MediaPlayer1.currentMedia

                        End If
                    End If
                End If


            Next
        End If
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Dim x, y As Integer

        x = Cursor.Position.X
        y = Cursor.Position.Y

        For Each ctrl As Control In Panel1.Controls
            Dim beginx As Integer = 0
            Dim endx As Integer = 0
            Dim beginy As Integer = 0
            Dim endy As Integer = 0
            beginx = ctrl.Left + Panel1.Left + Me.Left
            beginy = ctrl.Top + Panel1.Top + Me.Top
            endx = ctrl.Left + ctrl.Width + Panel1.Left + Me.Left
            endy = ctrl.Top + ctrl.Height + Panel1.Top + Me.Top

            If x >= beginx And x <= endx Then

                If y >= beginy And y <= endy Then
                    If (Not ctrl.Tag = "example") And (Not ctrl.Tag = "") Then
                        Dim fin As FileInfo = New FileInfo(ctrl.Tag)
                        Label13.Text = fin.Name
                        PictureBox2.Image = PictureBox2.Image.FromFile(((WorkingFolder & "\Thumbnails\" & fin.Name.Replace(fin.Extension, "") & "3.jpg").Replace("\\", "\")))
                        'MsgBox((Application.StartupPath & "\Thumbnails\" & fin.Name.Replace(fin.Extension, "") & "1.jpg").Replace("\\", "\"))
                        fin = Nothing
                        MediaPlayer1.URL = ctrl.Tag
                        MediaPlayer1.Ctlcontrols.play()
                        'DosShellCommand("""" & ctrl.Tag & """")
                        Exit For
                    End If

                End If
            End If


        Next
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        If Not MediaPlayer1.currentMedia Is Nothing Then
            MediaPlayer1.Ctlcontrols.play()
        End If
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        If Not MediaPlayer1.currentMedia Is Nothing Then
            MediaPlayer1.Ctlcontrols.stop()
        End If
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        If Not MediaPlayer1.currentMedia Is Nothing Then
            MediaPlayer1.Ctlcontrols.pause()
        End If
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        If Not MediaPlayer1.currentMedia Is Nothing Then
        MediaPlayer1.Ctlcontrols.currentPosition = MediaPlayer1.Ctlcontrols.currentPosition - 30
        End If
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        If Not MediaPlayer1.currentMedia Is Nothing Then
            MediaPlayer1.Ctlcontrols.currentPosition = MediaPlayer1.Ctlcontrols.currentPosition + 30
        End If
    End Sub

    Private Sub CurrentPositionTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CurrentPositionTimer.Tick
        If Not MediaPlayer1.currentMedia Is Nothing Then
            Label10.Text = (MediaPlayer1.Ctlcontrols.currentPositionString) & " / " & MediaPlayer1.currentMedia.durationString
        End If
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        If Not MediaPlayer1.currentMedia Is Nothing Then
            MediaPlayer1.Ctlcontrols.currentPosition = MediaPlayer1.Ctlcontrols.currentPosition - 240
        End If
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        If Not MediaPlayer1.currentMedia Is Nothing Then
            MediaPlayer1.Ctlcontrols.currentPosition = MediaPlayer1.Ctlcontrols.currentPosition + 240
        End If
    End Sub

    Private Sub Load_Details(ByVal inputname As String, Optional ByVal action As String = "stop")
        Try
            Dim finfo As FileInfo
            For Each ctrl As Control In Panel1.Controls
                If Not ctrl.Tag = "example" Then
                    finfo = New FileInfo(ctrl.Tag)
                    If finfo.Name.Replace(finfo.Extension, "") = inputname Then
                        'Panel1.Controls.Item(Panel1.Controls.IndexOf(ctrl) + 9).Select()
                        'Panel1.Controls.Item(Panel1.Controls.IndexOf(ctrl) + 9).Focus()
                        '     Panel1.Top = ctrl.Top
                        If Not CurrentListItem = ListBox1.SelectedIndex Then
                            If ListBox1.SelectedIndex > CurrentListItem Then
                                If Panel1.Controls.Count >= (Panel1.Controls.IndexOf(ctrl) + 18) + 1 Then
                                    Panel1.ScrollControlIntoView(Panel1.Controls.Item(Panel1.Controls.IndexOf(ctrl) + 18))
                                    'Panel1.Controls.Item(Panel1.Controls.IndexOf(ctrl) + 18).Focus()
                                Else
                                    Panel1.ScrollControlIntoView(Panel1.Controls.Item(Panel1.Controls.IndexOf(ctrl) + 9))
                                    'Panel1.Controls.Item(Panel1.Controls.IndexOf(ctrl) + 9).Focus()
                                End If
                            Else
                                
                                Panel1.ScrollControlIntoView(Panel1.Controls.Item(Panel1.Controls.IndexOf(ctrl) - 0))
                                'Panel1.Controls.Item(Panel1.Controls.IndexOf(ctrl) - 0).Focus()


                            End If

                        End If



                        Label13.Text = finfo.Name
                        PictureBox2.Image = PictureBox2.Image.FromFile(((WorkingFolder & "\Thumbnails\" & finfo.Name.Replace(finfo.Extension, "") & "3.jpg").Replace("\\", "\")))
                        MediaPlayer1.URL = ctrl.Tag
                        If action = "" Or action = "stop" Then
                            MediaPlayer1.Ctlcontrols.stop()
                        End If
                        If action = "play" Then
                            MediaPlayer1.Ctlcontrols.stop()
                            DosShellCommand("""" & ctrl.Tag & """")
                        End If
                        If action = "preview" Then
                            MediaPlayer1.Ctlcontrols.play()
                        End If
                        Exit For
                    End If
                    End If
            Next
            finfo = Nothing
        Catch ex As Exception
            Error_Handler(ex, "ListBox Clicked")
        End Try
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        If Not ListBox1.SelectedItem Is Nothing Then
            Load_Details(ListBox1.SelectedItem.ToString())
            CurrentListItem = ListBox1.SelectedIndex
        End If
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        If Not ListBox1.SelectedItem Is Nothing Then
            Load_Details(ListBox1.SelectedItem.ToString(), "play")
        End If
    End Sub



    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        If Not ListBox1.SelectedItem Is Nothing Then
            Load_Details(ListBox1.SelectedItem.ToString(), "preview")
        End If
    End Sub
End Class
