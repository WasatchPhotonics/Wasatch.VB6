VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WasatchNET VB6 Demo"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Event Log"
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   10215
      Begin VB.TextBox txtEventLog 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   240
         Width           =   9975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Controls"
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   10215
      Begin VB.TextBox TextIntegrationTimeMS 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Text            =   "100"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton btnStart 
         Caption         =   "Start"
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnInit 
         Caption         =   "Initialize"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Integration Time (MS)"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spectra"
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.PictureBox MSChart1 
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4035
         ScaleWidth      =   9915
         TabIndex        =   4
         Top             =   240
         Width           =   9975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim spectrometer As WasatchNET.spectrometer
Dim pixels As Long

Private Sub log(msg As String)
    txtEventLog.Text = txtEventLog.Text & vbNewLine & msg
End Sub

Private Sub btnInit_Click()
    Dim wrapper As New WasatchNET.DriverVBAWrapper
    Dim driver As WasatchNET.driver
    Set driver = wrapper.instance
    
    Dim numberOfSpectrometers As Integer
    numberOfSpectrometers = driver.openAllSpectrometers()
    If (numberOfSpectrometers <= 0) Then
        MsgBox "No spectrometers found"
        Return
    End If
    log "Found " & numberOfSpectrometers & " spectrometer"
    
    Set spectrometer = driver.getSpectrometer(0)
    ' pixels = spectrometer.pixels
    
    Dim wavelengths() As Double
    Dim wavenumbers() As Double
    Dim excitationNM As Single
    
    wavelengths = spectrometer.wavelengths
    wavenumbers = spectrometer.wavenumbers
    excitationNM = spectrometer.EEPROM.laserExcitationWavelengthNMFloat
    log "Excitation = " & excitationNM & "nm"
    log "Pixels = " & spectrometer.pixels
        
    Dim wavecalCoeffs() As Single
    wavecalCoeffs = spectrometer.EEPROM.wavecalCoeffs
End Sub
