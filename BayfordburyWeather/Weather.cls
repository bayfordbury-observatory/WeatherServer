VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Weather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const ALERT_TITLE As String = "BayfordburyWeather3.Weather"
Private Const MSG_START As String = "Property Get "
Private Const MSG_FIN As String = " is not implemented"

'Exposed methods and properties

Private Sub Class_Initialize()

    m_bConnected = False
    If Not g_bRunExecutable Then g_iReferenceCount = g_iReferenceCount + 1
    
End Sub

Private Sub Class_Terminate()

    If Not g_bRunExecutable Then
        If g_iReferenceCount > 1 Then
            g_iReferenceCount = g_iReferenceCount - 1
        Else
            Unload frmMain
        End If
    End If

End Sub

Private Sub CheckConnected()
    If Not m_bConnected Or Not frmMain.Visible Then _
        Err.Raise vbObjectError + &H400, "BayfordburyWeather3.Weather", "The weather server is not connected"
End Sub

' =================
' PUBLIC PROPERTIES
' =================

Public Property Get AmbientTemperature() As Single

    CheckConnected
    AmbientTemperature = temp
    
End Property

Public Property Get BarometricPressure() As Single
    
    CheckConnected
    BarometricPressure = pressure

End Property

Public Property Get Clouds() As Single

    CheckConnected
    If cloud = 1 Then
        Clouds = 0
    ElseIf cloud = 2 Then
        Clouds = 0.5
    Else
        Clouds = 1
    End If
    
End Property

Public Property Get Connected() As Boolean

    Connected = m_bConnected
    
End Property

Public Property Let Connected(ByVal newVal As Boolean)

    m_bConnected = newVal
    
End Property

Public Property Get DewPoint() As Single

    CheckConnected
    DewPoint = dewpt
    
End Property

Public Property Get InsideTemperature() As Single

    CheckConnected
    InsideTemperature = intemp
    
End Property

Public Property Get Name() As String

    Name = "Bayfordbury Weather Server 3"
    
End Property

Public Property Get Precipitation() As Boolean

    CheckConnected
    If rain > 2 Then
        Precipitation = True
    Else
        Precipitation = False
    End If
    
End Property

Public Property Get RelativeHumidity() As Single

    CheckConnected
    RelativeHumidity = humidity / 100
    
End Property

Public Property Get Safe() As Boolean

    CheckConnected
    Safe = IsSafe
   
End Property

Public Property Get WindDirection() As Single

    CheckConnected
    WindDirection = windir
    
End Property

Public Property Get WindVelocity() As Single

    CheckConnected
    WindVelocity = wind * 0.539957
    
End Property

' ==============
' PUBLIC METHODS
' ==============

Public Sub SetupDialog()

    MsgBox "There is nothing to set up.", (vbOKOnly + vbInformation), ALERT_TITLE
    
End Sub
