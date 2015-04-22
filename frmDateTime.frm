VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDateTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DateTime"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "frmDateTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timeClock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   4920
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin TabDlg.SSTab sstDateTime 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   9
      TabHeight       =   520
      BackColor       =   -2147483638
      TabCaption(0)   =   "Date && Time"
      TabPicture(0)   =   "frmDateTime.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTimeZone"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Time Zone"
      TabPicture(1)   =   "frmDateTime.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboZone"
      Tab(1).Control(1)=   "picMap"
      Tab(1).Control(2)=   "chkSaving"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "&Time"
         Height          =   3375
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   3255
         Begin VB.PictureBox picClock 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000008&
            Height          =   2505
            Left            =   240
            ScaleHeight     =   2445
            ScaleWidth      =   2685
            TabIndex        =   22
            Top             =   240
            Width           =   2745
         End
         Begin MSComCtl2.UpDown updSec 
            Height          =   375
            Left            =   1920
            TabIndex        =   21
            Top             =   2880
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            BuddyControl    =   "txtSec"
            BuddyDispid     =   196616
            OrigLeft        =   2400
            OrigTop         =   3000
            OrigRight       =   2640
            OrigBottom      =   3375
            Max             =   59
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updMin 
            Height          =   375
            Left            =   1335
            TabIndex        =   20
            Top             =   2880
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            BuddyControl    =   "txtMin"
            BuddyDispid     =   196617
            OrigLeft        =   1680
            OrigTop         =   3000
            OrigRight       =   1920
            OrigBottom      =   3375
            Max             =   59
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updHour 
            Height          =   375
            Left            =   720
            TabIndex        =   19
            Top             =   2880
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtHour"
            BuddyDispid     =   196618
            OrigLeft        =   960
            OrigTop         =   3000
            OrigRight       =   1200
            OrigBottom      =   3375
            Max             =   12
            Min             =   1
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.OptionButton optTime 
            Caption         =   "PM"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   16
            Top             =   3120
            Width           =   615
         End
         Begin VB.OptionButton optTime 
            Caption         =   "AM"
            Height          =   195
            Index           =   0
            Left            =   2160
            TabIndex        =   15
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtSec 
            Height          =   375
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   14
            Text            =   "ss"
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox txtMin 
            Height          =   375
            Left            =   960
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "mm"
            Top             =   2880
            Width           =   375
         End
         Begin VB.TextBox txtHour 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "11"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3076
               SubFormatType   =   0
            EndProperty
            Height          =   375
            Left            =   360
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "hh"
            Top             =   2880
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "&Date"
         Height          =   3375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   3135
         Begin MSComCtl2.UpDown updYear 
            Height          =   375
            Left            =   2760
            TabIndex        =   18
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1900
            BuddyControl    =   "txtYear"
            BuddyDispid     =   196620
            OrigLeft        =   2640
            OrigTop         =   480
            OrigRight       =   2880
            OrigBottom      =   855
            Max             =   2100
            Min             =   1900
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSACAL.Calendar calDay 
            Height          =   2175
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   2895
            _Version        =   524288
            _ExtentX        =   5106
            _ExtentY        =   3836
            _StockProps     =   1
            BackColor       =   -2147483639
            Year            =   2004
            Month           =   4
            Day             =   3
            DayLength       =   1
            MonthLength     =   0
            DayFontColor    =   0
            FirstDay        =   1
            GridCellEffect  =   0
            GridFontColor   =   10485760
            GridLinesColor  =   -2147483632
            ShowDateSelectors=   0   'False
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   0   'False
            ShowTitle       =   0   'False
            ShowVerticalGrid=   0   'False
            TitleFontColor  =   10485760
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtYear 
            Height          =   405
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cboMonth 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkSaving 
         Caption         =   "Automatically adjust clock for &daylight saving changes"
         Height          =   495
         Left            =   -74280
         TabIndex        =   6
         Top             =   3720
         Width           =   5535
      End
      Begin VB.PictureBox picMap 
         AutoRedraw      =   -1  'True
         Height          =   2775
         Left            =   -74280
         Picture         =   "frmDateTime.frx":0342
         ScaleHeight     =   2715
         ScaleWidth      =   5475
         TabIndex        =   5
         Top             =   960
         Width           =   5535
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   -74280
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label lblTimeZone 
         Caption         =   "lblTimeZone"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3960
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Writer : Karl Lam
'http://www.karl-lam.net
'Ex8 --- Date Time with moving clock
'################################################################
Dim hasSavingsTime(50) As Boolean
Dim timeOffSet(50) As String
Dim initYear, initMonth, initDay, initHour, initMin, initSec As Double
Dim currYear, currMonth, currDay, currHour, currMin, currSec As Double

Private Sub chkSaving_Click()
    cmdApply.Enabled = True
End Sub

Private Sub Form_Load()
    
    ' Init the TimeZone, Date & disable the Apply Botton
    Call initTime
    Call initZone
    Call initDate
    
    cmdApply.Enabled = False
    timeClock.Enabled = True
End Sub
Private Sub initZone()

' Set the time zone values

    timeOffSet(0) = "-720"
    timeOffSet(1) = "-660"
    timeOffSet(2) = "-600"
    timeOffSet(3) = "-540"
    timeOffSet(4) = "-480"
    timeOffSet(5) = "-420"
    timeOffSet(6) = "-420"
    timeOffSet(7) = "-360"
    timeOffSet(8) = "-360"
    timeOffSet(9) = "-360"
    timeOffSet(10) = "-300"
    timeOffSet(11) = "-300"
    timeOffSet(12) = "-300"
    timeOffSet(13) = "-240"
    timeOffSet(14) = "-240"
    timeOffSet(15) = "-210"
    timeOffSet(16) = "-180"
    timeOffSet(17) = "-180"
    timeOffSet(18) = "-120"
    timeOffSet(19) = "-060"
    timeOffSet(20) = "+000"
    timeOffSet(21) = "+000"
    timeOffSet(22) = "+060"
    timeOffSet(23) = "+060"
    timeOffSet(24) = "+060"
    timeOffSet(25) = "+120"
    timeOffSet(26) = "+120"
    timeOffSet(27) = "+120"
    timeOffSet(28) = "+120"
    timeOffSet(29) = "+120"
    timeOffSet(30) = "+180"
    timeOffSet(31) = "+180"
    timeOffSet(32) = "+180"
    timeOffSet(33) = "+240"
    timeOffSet(34) = "+270"
    timeOffSet(35) = "+300"
    timeOffSet(36) = "+330"
    timeOffSet(37) = "+360"
    timeOffSet(38) = "+420"
    timeOffSet(39) = "+480"
    timeOffSet(40) = "+480"
    timeOffSet(41) = "+540"
    timeOffSet(42) = "+570"
    timeOffSet(43) = "+570"
    timeOffSet(44) = "+600"
    timeOffSet(45) = "+600"
    timeOffSet(46) = "+600"
    timeOffSet(47) = "+660"
    timeOffSet(48) = "+720"
    timeOffSet(49) = "+720"
    
' #Set the time zone values


' Set the Time Zone saving

    For i = 0 To 49
        hasSavingsTime(i) = True
    Next i
    
    hasSavingsTime(0) = False  ' (GMT -12:00) Eniwetok, Kwajalein
    hasSavingsTime(1) = False  ' (GMT -11:00) Midway Island, Samoa
    hasSavingsTime(2) = False  ' (GMT -10:00) Hawaii
    hasSavingsTime(5) = False  ' (GMT -07:00) Arizona
    hasSavingsTime(8) = False  ' (GMT -06:00) Mexico City, Tegucigalpa
    hasSavingsTime(9) = False  ' (GMT -06:00) Saskatchewan
    hasSavingsTime(10) = False ' (GMT -05:00) Bogota, Lima
    hasSavingsTime(12) = False ' (GMT -05:00) Indiana (East)
    hasSavingsTime(14) = False ' (GMT -04:00) Caracas, La Paz
    hasSavingsTime(17) = False ' (GMT -03:00) Buenos Aires, Georgetown
    hasSavingsTime(21) = False ' (GMT +00:00) Monrovia, Casablanca
    hasSavingsTime(28) = False ' (GMT +02:00) Harare, Pretoria
    hasSavingsTime(30) = False ' (GMT +03:00) Baghdad, Kuwait, Nairobi, Riyadh
    hasSavingsTime(33) = False ' (GMT +04:00) Abu Dhabi, Muscat, Tbilisi
    hasSavingsTime(34) = False ' (GMT +04:30) Kabul
    hasSavingsTime(35) = False ' (GMT +05:00) Islamabad, Karachi, Ekaterinburg, Tashkent
    hasSavingsTime(36) = False ' (GMT +05:30) Bombay, Calcutta, Madras, New Delhi, Colombo
    hasSavingsTime(37) = False ' (GMT +06:00) Almaty, Dhaka
    hasSavingsTime(38) = False ' (GMT +07:00) Bangkok, Jakarta, Hanoi
    hasSavingsTime(40) = False ' (GMT +08:00) Hong Kong, Perth, Singapore, Taipei
    hasSavingsTime(41) = False ' (GMT +09:00) Tokyo, Osaka, Sapporo, Seoul, Yakutsk
    hasSavingsTime(43) = False ' (GMT +09:30) Darwin
    hasSavingsTime(45) = False ' (GMT +10:00) Guam, Port Moresby, Vladivostok
    hasSavingsTime(47) = False ' (GMT +11:00) Magadan, Solomon Is., New Caledonia
    hasSavingsTime(48) = False ' (GMT +12:00) Fiji, Kamchatka, Marshall Is.
    
' #Set the Time Zone saving

' Add Time Zone to check box

    cboZone.AddItem "(GMT -12:00) Eniwetok, Kwajalein"
    cboZone.AddItem "(GMT -11:00) Midway Island, Samoa"
    cboZone.AddItem "(GMT -10:00) Hawaii"
    cboZone.AddItem "(GMT -09:00) Alaska"
    cboZone.AddItem "(GMT -08:00) Pacific Time (US and Canada); Tijuana"
    cboZone.AddItem "(GMT -07:00) Arizona"
    cboZone.AddItem "(GMT -07:00) Mountain Time (US and Canada)"
    cboZone.AddItem "(GMT -06:00) Central Time (US and Canada)"
    cboZone.AddItem "(GMT -06:00) Mexico City, Tegucigalpa"
    cboZone.AddItem "(GMT -06:00) Saskatchewan"
    cboZone.AddItem "(GMT -05:00) Bogota, Lima"
    cboZone.AddItem "(GMT -05:00) Eastern Time (US and Canada)"
    cboZone.AddItem "(GMT -05:00) Indiana (East)"
    cboZone.AddItem "(GMT -04:00) Atlantic Time (Canada)"
    cboZone.AddItem "(GMT -04:00) Caracas, La Paz"
    cboZone.AddItem "(GMT -03:30) Newfoundland"
    cboZone.AddItem "(GMT -03:00) Brasilia"
    cboZone.AddItem "(GMT -03:00) Buenos Aires, Georgetown"
    cboZone.AddItem "(GMT -02:00) Mid-Atlantic"
    cboZone.AddItem "(GMT -01:00) Azores, Cape Verde Is."
    cboZone.AddItem "(GMT +00:00) Greenwich Mean Time; Dublin, Edinburgh, London, Lisbon"
    cboZone.AddItem "(GMT +00:00) Monrovia, Casablanca"
    cboZone.AddItem "(GMT +01:00) Berlin, Stockhold, Rome, Bern, Brussels, Vienna"
    cboZone.AddItem "(GMT +01:00) Paris, Madrid, Amsterdam"
    cboZone.AddItem "(GMT +01:00) Prage, Warsaw, Budapest"
    cboZone.AddItem "(GMT +02:00) Athens, Helsinki, Istanbul"
    cboZone.AddItem "(GMT +02:00) Cairo"
    cboZone.AddItem "(GMT +02:00) Eastern Europe"
    cboZone.AddItem "(GMT +02:00) Harare, Pretoria"
    cboZone.AddItem "(GMT +02:00) Israel"
    cboZone.AddItem "(GMT +03:00) Baghdad, Kuwait, Nairobi, Riyadh"
    cboZone.AddItem "(GMT +03:00) Moscow, St. Petersburgh, Kazan, Volgograd"
    cboZone.AddItem "(GMT +03:00) Tehran"
    cboZone.AddItem "(GMT +04:00) Abu Dhabi, Muscat, Tbilisi"
    cboZone.AddItem "(GMT +04:30) Kabul"
    cboZone.AddItem "(GMT +05:00) Islamabad, Karachi, Ekaterinburg, Tashkent"
    cboZone.AddItem "(GMT +05:30) Bombay, Calcutta, Madras, New Delhi, Colombo"
    cboZone.AddItem "(GMT +06:00) Almaty, Dhaka"
    cboZone.AddItem "(GMT +07:00) Bangkok, Jakarta, Hanoi"
    cboZone.AddItem "(GMT +08:00) Beijing, Chongqing, Urumqi"
    cboZone.AddItem "(GMT +08:00) Hong Kong, Perth, Singapore, Taipei"
    cboZone.AddItem "(GMT +09:00) Tokyo, Osaka, Sapporo, Seoul, Yakutsk"
    cboZone.AddItem "(GMT +09:30) Adelaide"
    cboZone.AddItem "(GMT +09:30) Darwin"
    cboZone.AddItem "(GMT +10:00) Brisbane, Melbourne, Sydney"
    cboZone.AddItem "(GMT +10:00) Guam, Port Moresby, Vladivostok"
    cboZone.AddItem "(GMT +10:00) Hobart"
    cboZone.AddItem "(GMT +11:00) Magadan, Solomon Is., New Caledonia"
    cboZone.AddItem "(GMT +12:00) Fiji, Kamchatka, Marshall Is."
    cboZone.AddItem "(GMT +12:00) Wellington, Auckland"
    
' #Add Time Zone to check box

    cboZone.ListIndex = 40 'Set time zone to Hong Kong ^^

End Sub
Private Sub initTime()

    ' Get init Time at startup
    initHour = Hour(Now)
    initMin = Minute(Now)
    initSec = Second(Now)
    
    ' Init current time
    currHour = initHour
    currMin = initMin
    currSec = initSec
    
    'Display the time

    If initHour = 0 Or initHour = 12 Then
    
        txtHour.Text = 12
        
    ElseIf initHour < 12 Then
        
        txtHour.Text = initHour
    
    Else
        
        txtHour.Text = initHour - 12
    
    End If
    
    If initHour < 12 Then
    
        optTime(0).Value = True
    Else
    
        optTime(1).Value = True
    End If
    
    txtMin.Text = initMin
    txtSec.Text = initSec
    
End Sub
Private Sub refreshTime()

    currSec = currSec + 1
    
    If currSec = 60 Then
        currSec = 0
        currMin = currMin + 1
    End If
    
    If currMin = 60 Then
        currMin = 0
        currHour = currHour + 1
    End If
    
    If currHour = 24 Then
        optTime(0).Value = True
        currHour = 0
        currDay = currDay + 1
        
        Call refreshDate
    End If
    
    'Display the time

    If currHour = 0 Or currHour = 12 Then
    
        txtHour.Text = 12
        
    ElseIf currHour < 12 Then
        
        txtHour.Text = currHour
    
    Else
        
        txtHour.Text = currHour - 12
    
    End If
    
    If currHour < 12 Then
    
        optTime(0).Value = True
    Else
    
        optTime(1).Value = True
    End If
        
    txtMin.Text = currMin
    txtSec.Text = currSec

End Sub
Private Sub initDate()

    ' Init the listbox
    cboMonth.AddItem "January"
    cboMonth.AddItem "February"
    cboMonth.AddItem "March"
    cboMonth.AddItem "April"
    cboMonth.AddItem "May"
    cboMonth.AddItem "June"
    cboMonth.AddItem "July"
    cboMonth.AddItem "Auguest"
    cboMonth.AddItem "September"
    cboMonth.AddItem "October"
    cboMonth.AddItem "November"
    cboMonth.AddItem "December"
    
    initYear = Year(Date)
    initMonth = Month(Date)
    initDay = Day(Date)
    
    currYear = initYear
    currMonth = initMonth
    currDay = initDay
    
    txtYear.Text = initYear
    cboMonth.ListIndex = initMonth - 1
    calDay.Day = initDay
End Sub
Private Sub refreshDate()
    txtYear.Text = currYear
    cboMonth.ListIndex = currMonth - 1
    calDay.Day = currDay
End Sub


Private Sub calDay_AfterUpdate()
        
    Call setDay
    Call setYear
    Call setMonth
    
    txtYear.Text = currYear
    cboMonth.ListIndex = currMonth - 1
End Sub

Private Sub calDay_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cboMonth_Click()

    
    calDay.Month = getMonth()
End Sub

Private Sub cboMonth_GotFocus()
    cmdApply.Enabled = True
End Sub

Private Sub cboZone_Click()

    ' Refresh the Map, checkbox
    cmdApply.Enabled = True
    chkSaving.Enabled = hasSavingsTime(cboZone.ListIndex)
    Call changeMap
    lblTimeZone.Caption = "Current time zone: " + Mid(cboZone.Text, 14)
End Sub

Private Sub cmdApply_Click()

    cmdApply.Enabled = False
    Call messShow(cmdApply)
    
    timeClock.Enabled = True
    cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()

    Call messShow(cmdCancel)
    Unload frmDateTime
End Sub

Private Sub cmdOK_Click()
    
    Call messShow(cmdOK)
    Unload frmDateTime
End Sub

Private Sub changeMap()

    ' Change the Time Zone to the center
    intMapOffSet = (Val(timeOffSet(cboZone.ListIndex)) / 4 * picMap.Width / 360) + picMap.Width / 2
    
    If intMapOffSet > 0 Then
    
        intMapOffSet = -(intMapOffSet)
    End If
    
    picMap.Cls
    picMap.PaintPicture picMap.Picture, intMapOffSet, 0
End Sub


Private Function errCheck(objIn As TextBox, key As Integer, min As Integer, max As Integer)

    ' Check the input Ascii code, return 0 when out of range
    If key >= 48 And key <= 57 Then
       
            strTemp1 = Mid(objIn.Text, 1, objIn.SelStart)
            strTemp2 = Chr(key)
            strTemp3 = Mid(objIn.Text, objIn.SelStart + objIn.SelLength + 1)
            intValue = strTemp1 + strTemp2 + strTemp3
                        
            If intValue >= min And intValue <= max Then
            
                errCheck = key
            Else
            
                errCheck = 0
            End If
            
    ElseIf key = 8 Or key = 35 Or key = 36 Or key = 37 Or key = 39 Then
    
        errCheck = key
    Else
    
        errCheck = 0
    End If
    
End Function


Private Sub optTime_GotFocus(Index As Integer)

    cmdApply.Enabled = True
    timeClock.Enabled = False
End Sub

Private Sub timeClock_Timer()

    Call refreshTime
    
End Sub

Private Sub txtHour_Change()

    Call drawClockFace(picClock, getHour(False), getMin(), getSec())
    
    Call setHour
End Sub

Private Sub txtHour_GotFocus()

    cmdApply.Enabled = True
    timeClock.Enabled = False

End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)

    KeyAscii = errCheck(txtHour, KeyAscii, 1, 12)
End Sub

Private Sub txtMin_Change()

    Call drawClockFace(picClock, getHour(False), getMin(), getSec())
    Call setMin
End Sub

Private Sub txtMin_GotFocus()
    cmdApply.Enabled = True
    timeClock.Enabled = False
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)

    KeyAscii = errCheck(txtMin, KeyAscii, 0, 59)
End Sub

Private Sub txtSec_Change()

    Call drawClockFace(picClock, getHour(False), getMin(), getSec())
    Call setSec
End Sub

Private Sub txtSec_GotFocus()
    cmdApply.Enabled = True
    timeClock.Enabled = False
End Sub

Private Sub txtSec_KeyPress(KeyAscii As Integer)

    KeyAscii = errCheck(txtSec, KeyAscii, 0, 59)
End Sub

Private Sub txtYear_Change()

    calDay.Year = getYear()
End Sub

Private Sub txtYear_GotFocus()

    cmdApply.Enabled = True
    
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)

    KeyAscii = errCheck(txtYear, KeyAscii, 1900, 2100)
End Sub

Public Sub drawClockFace(onObj As PictureBox, clockHour As Double, clockMin As Double, clockSec As Double)

    ' Clear the clock
    onObj.Cls

    Dim intMinute As Double
    
    intMinute = 0
    
    ' Set my own color
    Dim cFace, cHour, cMin, cSec As ColorConstants
    cFace = RGB(180, 255, 180)
    cHour = RGB(0, 255, 0)
    cMin = RGB(200, 200, 255)
    cSec = RGB(255, 0, 0)
    
    Dim smallDim As Integer           ' smaller of width and height of box
    Dim centerX As Integer            ' x for center of the clock
    Dim centerY As Integer            ' y for center of the clock
    Dim X(4) As Integer                  ' x position of dot
    Dim Y(4) As Integer
    Dim i As Integer
        
    centerX = onObj.Width / 2
    centerY = onObj.Height / 2
    smallDim = onObj.Width
    If onObj.Height < smallDim Then smallDim = onObj.Height
    
    
    ' Draw the minute dots
    For i = 0 To 59

        Const clockInset As Integer = 100 ' distance from clock edge to box edge
        Dim dotRadius As Integer
        If intMinute Mod 15 = 0 Then
        
            dotRadius = 45
            
        ElseIf intMinute Mod 5 = 0 Then
        
            dotRadius = 35
        Else
        
            dotRadius = 15
        End If
    
        ' calculate the center of the clock and its radius so it fits in the box
        
        Dim clockRadius As Integer        ' radius of the clock
        clockRadius = smallDim / 2 - clockInset

        ' calculate the position of the dot

        X(0) = centerX + clockX(intMinute, clockRadius - dotRadius)
        Y(0) = centerY + clockY(intMinute, clockRadius - dotRadius)
        
        ' draw a filled circle
        onObj.FillStyle = 0 ' filled primitives
        onObj.FillColor = cFace
        onObj.Circle (X(0), Y(0)), dotRadius, vbBlack
        
        intMinute = intMinute + 1
    Next i
    
    ' Draw hour hand
    Dim adjHour As Double
    adjHour = (clockHour + clockMin / 60) * 5
    
    X(0) = centerX + clockX(adjHour, clockRadius * 0.6)
    Y(0) = centerY + clockY(adjHour, clockRadius * 0.6)
    
    X(2) = centerX + clockX(adjHour + 30, clockRadius * 0.15)
    Y(2) = centerY + clockY(adjHour + 30, clockRadius * 0.15)
     
    
    For i = 0 To 100
    
        X(1) = centerX + clockX(adjHour + 15, clockRadius * 0.001 * i)
        Y(1) = centerY + clockY(adjHour + 15, clockRadius * 0.001 * i)
        
        onObj.Line (X(0), Y(0))-(X(1), Y(1)), cHour
        onObj.Line (X(2), Y(2))-(X(1), Y(1)), cHour
        
        X(3) = centerX + clockX(adjHour + 45, clockRadius * 0.001 * i)
        Y(3) = centerY + clockY(adjHour + 45, clockRadius * 0.001 * i)
        
        onObj.Line (X(0), Y(0))-(X(3), Y(3)), cHour
        onObj.Line (X(2), Y(2))-(X(3), Y(3)), cHour
        
    Next i
    
    'Draw minute hand
    Dim adjMin As Double
    adjMin = (clockMin + clockSec / 60)
    
    X(0) = centerX + clockX(adjMin, clockRadius * 0.8)
    Y(0) = centerY + clockY(adjMin, clockRadius * 0.8)
    
    X(2) = centerX + clockX(adjMin + 30, clockRadius * 0.2)
    Y(2) = centerY + clockY(adjMin + 30, clockRadius * 0.2)
    
        
    For i = 0 To 60
    
        X(1) = centerX + clockX(adjMin + 15, clockRadius * 0.001 * i)
        Y(1) = centerY + clockY(adjMin + 15, clockRadius * 0.001 * i)
        
        onObj.Line (X(0), Y(0))-(X(1), Y(1)), cMin
        onObj.Line (X(2), Y(2))-(X(1), Y(1)), cMin
        
        X(3) = centerX + clockX(adjMin + 45, clockRadius * 0.001 * i)
        Y(3) = centerY + clockY(adjMin + 45, clockRadius * 0.001 * i)
        
        onObj.Line (X(0), Y(0))-(X(3), Y(3)), cMin
        onObj.Line (X(2), Y(2))-(X(3), Y(3)), cMin
            
    Next i
            
    
    'Draw second hand
    X(0) = centerX + clockX(clockSec, clockRadius * 0.9)
    Y(0) = centerY + clockY(clockSec, clockRadius * 0.9)
    
    X(2) = centerX + clockX(clockSec + 30, clockRadius * 0.4)
    Y(2) = centerY + clockY(clockSec + 30, clockRadius * 0.4)
    
    
    For i = 0 To 30
        X(1) = centerX + clockX(clockSec + 15, clockRadius * 0.001 * i)
        Y(1) = centerY + clockY(clockSec + 15, clockRadius * 0.001 * i)
        
        onObj.Line (X(0), Y(0))-(X(1), Y(1)), cSec
        onObj.Line (X(2), Y(2))-(X(1), Y(1)), cSec
    
        X(3) = centerX + clockX(clockSec + 45, clockRadius * 0.001 * i)
        Y(3) = centerY + clockY(clockSec + 45, clockRadius * 0.001 * i)
        
        onObj.Line (X(0), Y(0))-(X(3), Y(3)), cSec
        onObj.Line (X(2), Y(2))-(X(3), Y(3)), cSec
 
    Next i
    
End Sub

Public Function clockX(minuteVal As Double, radius As Integer)
    Const PI As Double = 3.14159265
    Dim angle As Double

    angle = (PI * 2 * minuteVal) / 60 - PI / 2
    clockX = CInt(radius * Cos(angle))
End Function

Public Function clockY(minuteVal As Double, radius As Integer)
    Dim angle As Double
    Dim PI As Double
    PI = 3.14159265

    angle = (PI * 2 * minuteVal) / 60 - PI / 2
    clockY = CInt(radius * Sin(angle))
End Function

Private Function getHour(hourType As Boolean)
    
    ' Return the true time (24 hours)
    If hourType Then
    
        If optTime(0).Value Then
        
            If Val(txtHour.Text) = 12 Then
                getHour = 12
            Else
                getHour = Val(txtHour.Text)
            End If
            
        Else
        
            If Val(txtHour.Text) = 12 Then
                getHour = 12
            Else
                getHour = Val(txtHour.Text) + 12
            End If
        End If
    
    ' Return the false (12 hours)
    Else
    
        getHour = Val(txtHour.Text)
    End If

End Function

Private Function getMin()

    getMin = Val(txtMin.Text)
End Function

Private Function getSec()

    getSec = Val(txtSec.Text)
End Function

Private Function getYear()

    getYear = Val(txtYear.Text)
End Function
Private Function getMonth()

    getMonth = cboMonth.ListIndex + 1
End Function
Private Function getDay()

    getDay = calDay.Day
End Function

Private Sub updHour_Change()

    cmdApply.Enabled = True
    timeClock.Enabled = False
End Sub

Private Sub updMin_Change()

    cmdApply.Enabled = True
    timeClock.Enabled = False
End Sub

Private Sub updSec_Change()

    cmdApply.Enabled = True
    timeClock.Enabled = False
End Sub

Private Sub updYear_Change()

    cmdApply.Enabled = True
    
End Sub
Private Sub messShow(objIn As CommandButton)
    
    ' Show the Year, Month, Day, Hour, Minute, Second
    Dim blnDayLight As Boolean
    blnDayLight = False
    
    If hasSavingsTime(cboZone.ListIndex) And chkSaving.Value = 1 Then
    
        blnDayLight = True
    End If
    
    Dim strStart, strEnd, strMess As String
    
    ' Make a string for the Button
    If objIn.Name = "cmdOK" Then
    
        strStart = "Confirmed..."
        strEnd = ""
    ElseIf objIn.Name = "cmdApply" Then
    
        strStart = "Appied..."
        strEnd = ""
    Else
        strStart = "Canceled..."
        strEnd = "not "
    End If
    
    ' Make a string for the Values
    strMess = strStart + vbCrLf + _
            "===================================" + vbCrLf + _
            "Year = " + Str(getYear()) + vbCrLf + _
            "Month = " + Str(getMonth()) + vbCrLf + _
            "Day = " + Str(getDay()) + vbCrLf + _
            "Hour = " + Str(getHour(True)) + vbCrLf + _
            "Minute = " + Str(getMin()) + vbCrLf + _
            "Second = " + Str(getSec()) + vbCrLf + _
            "Timezone = " + cboZone.Text + "[#" + Str(cboZone.ListIndex) + "]" + vbCrLf + _
            "Auto Daylight = " + Str(blnDayLight) + vbCrLf + _
            "===================================" + vbCrLf + _
            "(Time " + strEnd + "saved)"
    
    intMess = MsgBox(strMess, vbOKOnly, "Date and Time")
    
End Sub
Private Sub setHour()

    If optTime(0).Value Then
        
        If Val(txtHour.Text) = 12 Then
        
                currHour = 0
        Else
                currHour = Val(txtHour.Text)
        End If
            
    Else
    
        If Val(txtHour.Text) = 12 Then
            currHour = 12
        Else
            currHour = Val(txtHour.Text) + 12
        End If
        
    End If
    
    
End Sub
Private Sub setMin()
    currMin = Val(txtMin.Text)
End Sub
Private Sub setSec()

    currSec = Val(txtSec.Text)

End Sub
Private Sub setYear()
    
    currYear = calDay.Year
End Sub

Private Sub setMonth()
    
    currMonth = calDay.Month
End Sub

Private Sub setDay()
    
    currDay = calDay.Day
End Sub
