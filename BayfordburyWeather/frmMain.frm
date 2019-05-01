VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bayfordbury Weather Server"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6510
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CheckBox DayOK 
      BackColor       =   &H00400000&
      Caption         =   "Daylight is safe"
      ForeColor       =   &H000080FF&
      Height          =   250
      Left            =   3000
      TabIndex        =   16
      Top             =   6915
      Width           =   2055
   End
   Begin VB.CheckBox CloudsOK 
      BackColor       =   &H00400000&
      Caption         =   "Clouds are safe"
      ForeColor       =   &H000080FF&
      Height          =   250
      Left            =   3000
      TabIndex        =   15
      Top             =   6615
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   2100
      Left            =   1
      Top             =   0
   End
   Begin VB.Label Label25 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   5040
      TabIndex        =   58
      Top             =   7180
      Width           =   1500
   End
   Begin VB.Label Label40 
      BackColor       =   &H00400000&
      Caption         =   "Last v. cloudy:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   56
      Top             =   5720
      Width           =   1100
   End
   Begin VB.Label Label39 
      BackColor       =   &H00400000&
      Caption         =   "Last gust:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   55
      Top             =   4220
      Width           =   1100
   End
   Begin VB.Label Label38 
      BackColor       =   &H00400000&
      Caption         =   "Last wet:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   54
      Top             =   4520
      Width           =   975
   End
   Begin VB.Label Label37 
      BackColor       =   &H00400000&
      Caption         =   "Last rain:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   53
      Top             =   4820
      Width           =   1005
   End
   Begin VB.Label Label36 
      BackColor       =   &H00400000&
      Caption         =   "Last cloudy:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   52
      Top             =   5420
      Width           =   1100
   End
   Begin VB.Label Label35 
      BackColor       =   &H00400000&
      Caption         =   "Last light:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   51
      Top             =   5120
      Width           =   1005
   End
   Begin VB.Label lastGustLbl 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   50
      Top             =   4220
      Width           =   1200
   End
   Begin VB.Label lastWetLbl 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   49
      Top             =   4520
      Width           =   1200
   End
   Begin VB.Label lastRainLbl 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   48
      Top             =   4820
      Width           =   1200
   End
   Begin VB.Label lastCloudyLbl 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   47
      Top             =   5420
      Width           =   1200
   End
   Begin VB.Label lastLightLbl 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   46
      Top             =   5120
      Width           =   1200
   End
   Begin VB.Label lastVeryCloudyLbl 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   45
      Top             =   5720
      Width           =   1200
   End
   Begin VB.Label Label28 
      BackColor       =   &H00400000&
      Caption         =   "Weather"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   240
      TabIndex        =   44
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label26 
      BackColor       =   &H00400000&
      Caption         =   "Solar radiation:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   200
      TabIndex        =   43
      Top             =   2300
      Width           =   1100
   End
   Begin VB.Label txtDaylight 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   1400
      TabIndex        =   42
      Top             =   2300
      Width           =   1200
   End
   Begin VB.Label Label24 
      BackColor       =   &H00400000&
      Caption         =   "Last reading: "
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3040
      TabIndex        =   41
      Top             =   2800
      Width           =   2400
   End
   Begin VB.Label Label23 
      BackColor       =   &H00400000&
      Caption         =   "Last reading: "
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   2800
      Width           =   2400
   End
   Begin VB.Label Label22 
      BackColor       =   &H00400000&
      Caption         =   "Bad weather"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   39
      Top             =   3875
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   38
      Top             =   2300
      Width           =   1125
   End
   Begin VB.Label Label20 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   37
      Top             =   2000
      Width           =   1125
   End
   Begin VB.Label Label19 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   36
      Top             =   1700
      Width           =   1125
   End
   Begin VB.Label Label18 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   35
      Top             =   1400
      Width           =   1125
   End
   Begin VB.Label Label17 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   34
      Top             =   1100
      Width           =   1125
   End
   Begin VB.Label Label16 
      BackColor       =   &H00400000&
      Caption         =   "5min Max:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   33
      Top             =   1700
      Width           =   1100
   End
   Begin VB.Label Label15 
      BackColor       =   &H00400000&
      Caption         =   "5min Min:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   32
      Top             =   2000
      Width           =   1100
   End
   Begin VB.Label Label14 
      BackColor       =   &H00400000&
      Caption         =   "5min Diff:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   31
      Top             =   2300
      Width           =   1100
   End
   Begin VB.Label Label13 
      BackColor       =   &H00400000&
      Caption         =   "Sky temperature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   30
      Top             =   200
      Width           =   1500
   End
   Begin VB.Label Label12 
      BackColor       =   &H00400000&
      Caption         =   "Zenith:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   29
      Top             =   800
      Width           =   1100
   End
   Begin VB.Label Label11 
      BackColor       =   &H00400000&
      Caption         =   "5min av:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   28
      Top             =   1100
      Width           =   1100
   End
   Begin VB.Label Label10 
      BackColor       =   &H00400000&
      Caption         =   "5min sdtev:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   27
      Top             =   1400
      Width           =   1100
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   26
      Top             =   800
      Width           =   1125
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      X1              =   2700
      X2              =   2700
      Y1              =   100
      Y2              =   7120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   1
      X1              =   0
      X2              =   2640
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      X1              =   0
      X2              =   2640
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label txtDelay 
      BackColor       =   &H00400000&
      Caption         =   "Last reading: "
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   6000
      Width           =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      Index           =   0
      X1              =   100
      X2              =   5400
      Y1              =   3200
      Y2              =   3200
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "Sky brightness (mag/""^2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   3400
      Width           =   2400
   End
   Begin VB.Label txtVisual 
      BackColor       =   &H00400000&
      Caption         =   "Visual"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   5300
      Width           =   1000
   End
   Begin VB.Label txtClear 
      BackColor       =   &H00400000&
      Caption         =   "Clear"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5000
      Width           =   1000
   End
   Begin VB.Label txtB 
      BackColor       =   &H00400000&
      Caption         =   "B"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4700
      Width           =   735
   End
   Begin VB.Label txtV 
      BackColor       =   &H00400000&
      Caption         =   "V"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4400
      Width           =   735
   End
   Begin VB.Label txtR 
      BackColor       =   &H00400000&
      Caption         =   "R"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4100
      Width           =   735
   End
   Begin VB.Label txtI 
      BackColor       =   &H00400000&
      Caption         =   "I"
      ForeColor       =   &H000080FF&
      Height          =   250
      Left            =   240
      TabIndex        =   18
      Top             =   3800
      Width           =   735
   End
   Begin VB.Label Version 
      BackColor       =   &H00400000&
      Caption         =   "2016-08-25    V3.1"
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   5040
      TabIndex        =   17
      Top             =   6960
      Width           =   1500
   End
   Begin VB.Label txtRain 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   1400
      TabIndex        =   14
      Top             =   2000
      Width           =   1000
   End
   Begin VB.Label txtWindVelocity 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   1400
      TabIndex        =   13
      Top             =   1400
      Width           =   1000
   End
   Begin VB.Label txtWindDirection 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   1400
      TabIndex        =   12
      Top             =   1700
      Width           =   1000
   End
   Begin VB.Label txtRelativeHumidity 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   1400
      TabIndex        =   11
      Top             =   1100
      Width           =   1000
   End
   Begin VB.Label txtDewPoint 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   1400
      TabIndex        =   10
      Top             =   800
      Width           =   1000
   End
   Begin VB.Label txtclouds 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   4200
      TabIndex        =   9
      Top             =   500
      Width           =   1500
   End
   Begin VB.Label txtAmbientTemperature 
      BackColor       =   &H00400000&
      Caption         =   "-"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   1400
      TabIndex        =   8
      Top             =   500
      Width           =   1000
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400000&
      Caption         =   "Wind Speed:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   200
      TabIndex        =   7
      Top             =   1400
      Width           =   1000
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Caption         =   "Wind Direction:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   200
      TabIndex        =   6
      Top             =   1700
      Width           =   1100
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Caption         =   "Humidity:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   200
      TabIndex        =   5
      Top             =   1100
      Width           =   1000
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      Caption         =   "Dew Point:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   200
      TabIndex        =   4
      Top             =   800
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400000&
      Caption         =   "Wide angle:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   3000
      TabIndex        =   3
      Top             =   500
      Width           =   1000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Temperature:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   200
      TabIndex        =   2
      Top             =   500
      Width           =   1100
   End
   Begin VB.Label SafeLabel 
      BackColor       =   &H00400000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   800
      Left            =   2900
      TabIndex        =   1
      Top             =   3400
      Width           =   2600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   "Rain:"
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   200
      TabIndex        =   0
      Top             =   2000
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Timer1_Timer()
On Error GoTo ProcError

        Dim successful, AllSafe, tooRainy, tooCloudy, tooBright, tooWindy, tooCold, tooHot, tooHumid, tooOld, preSafe As Boolean
        Dim unsafeWind, unsafeCloud, unsafeVeryCloud, unsafeRain, unsafeWet, unsafeLight, unsafeTemp, unsafeHum, unsafeClose, unsafeData, safeLevel, skyCondition As Integer
        
        Dim magi, magr, magv, magb, magc, magw, mags As Single
        
        Dim rad5sAv, rad5mAv, rad5mStdev, rad5mMax, rad5mMin, rad5mDiff As Single
          
        ref.read
        
        Dim reason As String
        
        Dim error As String
        
        successful = ref.successful
        
        If successful = True Then
        
            Label25.Caption = "Library v" + ref.Version
            
            'Text1.Text = CStr(ref.length) + " " + ref.Error
              
            windir = ref.wdirD
            wind = ref.windD
           
            SkyTemperature = ref.skytD
            intemp = 0
            pressure = 0
            dewpt = ref.dewD
            humidity = ref.humD
            SolarRad = ref.radD
            
            magi = ref.magiD
            magr = ref.magrD
            magv = ref.magvD
            magb = ref.magbD
            magc = ref.magcD
            mags = ref.magsD
            
            unsafeWind = ref.unsafeWind
            unsafeCloud = ref.unsafeCloud
            unsafeVeryCloud = ref.unsafeVeryCloud
            unsafeRain = ref.unsafeRain
            unsafeWet = ref.unsafeWet
            unsafeLight = ref.unsafeLight
            unsafeTemp = ref.unsafeTemp
            unsafeHum = ref.unsafeHum
            unsafeClose = ref.unsafeClose
            unsafeData = ref.unsafeData
            safeLevel = ref.safeLevel
            skyCondition = ref.skyCondition

    
            rain = ref.rainB
            cloud = 0
            
            txtRain.Caption = rain
            
             
            
            temp = ref.tempD
            
            txtWindVelocity.Caption = CStr(wind) + " kph "
            txtRelativeHumidity.Caption = CStr(humidity) + " " & Chr(37)
            txtDewPoint.Caption = CStr(dewpt) & Chr(176) + "C"
            'txtBarometricPressure.caption = pressure + " hPa"
            txtAmbientTemperature.Caption = CStr(temp) & Chr(176) + "C"
            'txtSkyTemperature.Caption = SkyTemperature & Chr(176) + "C"
            txtWindDirection.Caption = CStr(windir) & Chr(176)
            txtDaylight.Caption = CStr(SolarRad) + " W/m^2"
            
            If ref.radLR > 30 Then
                Label4.Caption = " - " & Chr(176) + "C"
                Label17.Caption = " - " & Chr(176) + "C"
                Label18.Caption = " - " & Chr(176) + "C"
                Label19.Caption = " - " & Chr(176) + "C"
                Label20.Caption = " - " & Chr(176) + "C"
                Label21.Caption = " - " & Chr(176) + "C"
                 Else
                Label4.Caption = ref.zenD & Chr(176) + "C"
                Label17.Caption = ref.zenavD & Chr(176) + "C"
                Label18.Caption = ref.zenstdevD & Chr(176) + "C"
                Label19.Caption = ref.zenmaxD & Chr(176) + "C"
                Label20.Caption = ref.zenminD & Chr(176) + "C"
                Label21.Caption = ref.zendifD & Chr(176) + "C"
              
            End If
                    
             If SkyTemperature = -998 Then
                txtclouds.Caption = "Wet, can't read"
            ElseIf SkyTemperature < -21 Then
                txtclouds.Caption = CStr(SkyTemperature) & Chr(176) + "C (Clear)"
            ElseIf SkyTemperature < -10 Then
                txtclouds.Caption = CStr(SkyTemperature) & Chr(176) + "C (Cloudy)"
            Else
                txtclouds.Caption = CStr(SkyTemperature) & Chr(176) + "C (Very Cloudy)"
            End If
            
            
             
            If magi > 8.1 And ref.photometerLR < 120 Then
               txtI.Caption = "I: " + CStr(magi)
            Else
               txtI.Caption = "I: -"
            End If
            
            If magr > 8.1 And ref.photometerLR < 120 Then
               txtR.Caption = "R: " + CStr(magr)
            Else
               txtR.Caption = "R: -"
            End If
            
            If magv > 8.1 And ref.photometerLR < 120 Then
               txtV.Caption = "V: " + CStr(magv)
            Else
               txtV.Caption = "V: -"
            End If
            
            If magb > 8.1 And ref.photometerLR < 120 Then
               txtB.Caption = "B: " + CStr(magb)
            Else
              txtB.Caption = "B: -"
            End If
            
            If magc > 8.1 And ref.photometerLR < 120 Then
               txtClear.Caption = "Clear: " + CStr(magc)
            Else
               txtClear.Caption = "Clear: -"
            End If
            
            If mags > 0 And ref.photometerLR < 60 Then
               txtVisual.Caption = "Visual: " + CStr(mags)
            Else
               txtVisual.Caption = "Visual: -"
            End If
            
            txtDelay.Caption = "Last reading: " + ref.timeUnitsSecs(ref.photometerLR)
            Label23.Caption = "Last reading: " + ref.timeUnitsSecs(ref.weatherLR)
            Label24.Caption = "Last reading: " + ref.timeUnitsSecs(ref.cloudLR)
                                 
            lastGustLbl.Caption = ref.timeUnits(ref.gustTime)
            lastWetLbl.Caption = ref.timeUnits(ref.wetTime)
            lastRainLbl.Caption = ref.timeUnits(ref.rainTime)
            lastLightLbl.Caption = ref.timeUnits(ref.lightTime)
            lastCloudyLbl.Caption = ref.timeUnits(ref.cloudyTime)
            lastVeryCloudyLbl.Caption = ref.timeUnits(ref.veryCloudyTime)
            
            'If tooRainy Or tooCloudy Or tooBright Or tooWindy Or tooHot Or tooCold Or tooHumid Or tooOld Then
            preSafe = True
            
            reason = "none"
            
            If unsafeClose > 0 Then
                'unsafe
                If DayOK.Value = 1 And CloudsOK.Value = 1 Then
                    'override
                Else
                    preSafe = False
                    reason = "hardware closure"
                End If
            
            Else
                'ok
                
            End If
            
            If unsafeCloud > 0 Then
                If unsafeCloud > 4 Then
                    'unsafe
                    If CloudsOK.Value = 1 Then
                        'override
                    Else
                        preSafe = False
                        reason = "cloud"
                    End If
                    
                    txtclouds.ForeColor = &HFF&
                Else
                    'striking
                    txtclouds.ForeColor = &H80FF&
                End If
            Else
                'ok
                txtclouds.ForeColor = &HFF00&
            End If
            
             If unsafeTemp > 0 Then
                    'unsafe
                    preSafe = False
                    reason = "temperature"
                    txtAmbientTemperature.ForeColor = &HFF&
             Else
                'ok
                txtAmbientTemperature.ForeColor = &HFF00&
            End If
            
            If unsafeHum > 0 Then
                    'unsafe
                    preSafe = False
                    reason = "humidity"
                    txtRelativeHumidity.ForeColor = &HFF&
            Else
                'ok
                txtRelativeHumidity.ForeColor = &HFF00&
            End If
            
            If unsafeWind > 0 Then
                If unsafeWind > 4 Then
                    'unsafe
                    preSafe = False
                    reason = "wind"
                    txtWindVelocity.ForeColor = &HFF&
                Else
                    'striking
                    txtWindVelocity.ForeColor = &H80FF&
                End If
            Else
                'ok
                txtWindVelocity.ForeColor = &HFF00&
            End If
            
            If unsafeLight > 0 Then
                If unsafeLight > 4 Then
                    'unsafe
                    If DayOK.Value = 1 Then
                        'override
                    Else
                        preSafe = False
                        reason = "light"
                    End If
                    txtDaylight.ForeColor = &HFF&
                Else
                    'striking
                    txtDaylight.ForeColor = &H80FF&
                End If
            Else
                'ok
                txtDaylight.ForeColor = &HFF00&
            End If
            
            If unsafeRain > 5 Then
                If unsafeRain > 4 Then
                    'unsafe
                    preSafe = False
                    reason = "rain"
                    txtRain.ForeColor = &HFF&
                Else
                    'striking
                    txtRain.ForeColor = &H80FF&
                End If
            Else
                'ok
                txtRain.ForeColor = &HFF00&
            End If
            
            
            
            If unsafeData > 0 Then
                'unsafe
                preSafe = False
                reason = "data error"
            Else
                'ok
                
            End If
            
            
            If preSafe Then
                sinceunsafe = sinceunsafe + 1
                errorcount = 0
                daylight = 0
                unsafecount = 0
            Else
                 unsafecount = 5
            End If
            
            
            
            'Text1.Text = ref.dataout
            
        Else
            'Data could not be read
            
            unsafecount = unsafecount + 1
            
            reason = "read error"
            
            error = ref.error
            
            Text1.Text = error
        
        End If
                           
       'IsSafe = False
        If startcount < 2 Then
            startcount = startcount + 1
            
           IsSafe = True
           SafeLabel.ForeColor = &H80FF&
           SafeLabel.Caption = "Initialising"
            
       ElseIf unsafecount > 4 Then
       
           IsSafe = False
           sinceunsafe = 0
           unsafecount = 10
           SafeLabel.ForeColor = &HFF&
           SafeLabel.Caption = "UNSAFE due to " + reason
           
       ElseIf unsafecount > 0 Then
       
           IsSafe = True
           SafeLabel.ForeColor = &H80FF&
           SafeLabel.Caption = "Bad weather detected" & vbCrLf & "Strike " + CStr(unsafecount)
           
       ElseIf sinceunsafe < 10 Then
       
           IsSafe = False
           SafeLabel.ForeColor = &H80FF&
           sinceunsafe = sinceunsafe + CInt(Math.Round(Timer1.Interval / 1000))
           SafeLabel.Caption = "Safe, but unsafe too recently" & vbCrLf & "(" + CStr(10 - sinceunsafe) + "s)"
       
       Else
       
           IsSafe = True
           SafeLabel.ForeColor = &HFF00&
           SafeLabel.Caption = "Weather SAFE for observing"
           'If sinceunsafe  > 0 Then
           '    unsafe CInt(Math.Round(Timer1.Interval / 1000))
           
                        
       End If
    
       errorcount = 0
             
        
ProcExit:
  Exit Sub

ProcError:
        'txtWindVelocity.Caption = "-"
        'txtclouds.Caption = "-"
        'txtRelativeHumidity.Caption = "-"
        'txtDewPoint.Caption = "-"
        'txtAmbientTemperature.Caption = "-"
        'txtWindDirection.Caption = "-"
        'txtRain.Caption = "-"
    
    If errorcount < 2 Then
        errorcount = errorcount + 1
        If sinceunsafe < 30 Then
            SafeLabel.Caption = "ERROR: " + Err.Description + " Strike " + CStr(errorcount)
            IsSafe = False
            SafeLabel.ForeColor = &HFF&
        Else
            IsSafe = True
            SafeLabel.Caption = "ERROR: " + Err.Description + " Strike " + CStr(errorcount)
            SafeLabel.ForeColor = &H80FF&
        End If
    Else
        Text1.Text = "ERROR: " + Err.Description
        IsSafe = False
        SafeLabel.ForeColor = &HFF&
   End If
   
  Resume ProcExit
        
End Sub



