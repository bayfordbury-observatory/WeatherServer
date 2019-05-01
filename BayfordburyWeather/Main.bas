Attribute VB_Name = "ModMain"
Option Explicit
'Public variables
'Singles
Public windir, wind, temp, SkyTemperature, intemp, pressure, dewpt, humidity, SolarRad As Single
'Integers
Public rain, cloud, sinceunsafe, errorcount, daylight, unsafecount, startcount As Integer

'Bools
Public IsSafe, startup As Boolean
'Connected
Public m_bConnected As Boolean
Public g_iReferenceCount As Integer
Public g_bRunExecutable As Boolean
Public Util
Public Console


Public ref As BayfordburyWeatherClass.Weather
 
Sub Main()

   
        'Set Console = CreateObject("ACP.Console")
        'Set Util = CreateObject("ACP.Util")
        'Util.WeatherConnected = False
         'Console.PrintLine ("Weather connected")
         
        'Set initial params
        IsSafe = True
        sinceunsafe = 30
        unsafecount = 0
        startcount = 0
        startup = True
        
        m_bConnected = True
        g_iReferenceCount = 0
        g_bRunExecutable = (App.StartMode = vbSModeStandalone)
        
        Load frmMain
        
        'If g_bRunExecutable Then
        '    frmMain.WindowState = vbNormal
        'Else
        '    frmMain.WindowState = vbMinimized
        'End If
        
        frmMain.Show
        
        
        Set ref = New BayfordburyWeatherClass.Weather
        
        'Start timer
        Call frmMain.Timer1_Timer
        'frmMain.Timer1_Timer()

End Sub




