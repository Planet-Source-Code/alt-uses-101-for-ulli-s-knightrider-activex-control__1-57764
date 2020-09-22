VERSION 5.00
Begin VB.UserControl KnightRider 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   FillColor       =   &H0000FF00&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   MaskColor       =   &H00000000&
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
   Begin VB.PictureBox picBlend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   180
      Left            =   -585
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   86
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   1320
   End
End
Attribute VB_Name = "KnightRider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Ulli's KnightRider.ctl
'**************************************************************************************************
' Title: KnightRider ActiveX Custom Control
' Description: This is a KnightRider Custom Control with customizable Back- and ForeColors, Size,
' Speed, and Effect. The appearance and effects can be viewed in the IDE; that is it is already '
' active in Design Mode :- just set the Enabled-property to True. It was inspired by a previous
' submission to PSC.
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=50306&lngWId=1
'**************************************************************************************************
' Revised 12/17/04 by AlT.  Why you ask?  Well, I downloaded this some time ago and thought it
' was a neat effect but what to use it for?  Well, I found a little time to kill and went through
' my downloaded code.  I re-discovered this and decided to see if there was a fun way to make
' use of it.  I found a couple and decided to resubmit this to PSC to see what some of you evil
' geniuses could do with it.  If you got a little time to play, let's see what you got. ;-)
'
' My revisions:
' Added Paul Caton's timer class and a reference to the WinSubHook2 typelib included in the zip.
' In my opinion, Paul's code is one the most useful ever submitted to PSC.  I have found a
' multitude of uses for it and it is simply awesome coding.  His original WinSubHook submission
' can be found here:
'
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=51403&lngWId=1
'
' Also, I formatted the code and changed variable names to my taste.  Forgive me Ulli.  Finally,
' I added a TripComplete event for use in my demo.
'
' To use, make sure you have a reference to the WinSubHook2 typelib.  If for some reason you
' can not see the effect, substitute the use of the Alphablend (msimg32.dll) API for the use of
' the GDIAlphablend (gdi32.dll) API call.  The parameters are identical so you should only have
' to change the function name, uncomment the Alphablend api call, and comment the GDIAlphablend
' API call.  Votes are not necessary but if you vote, they are Ulli's.....
'**************************************************************************************************
Option Explicit
'**************************************************************************************************
' Enum Case Protection
'**************************************************************************************************
#If False Then
     Private LeftToRight As Long
     Private RightToLeft As Long
     Private Oscillating As Long
#End If

'**************************************************************************************************
' KnightRider Enums\Structs
'**************************************************************************************************
Public Enum eEffect
     LeftToRight
     RightToLeft
     Oscillating
End Enum ' eEffect

'**************************************************************************************************
' KnightRider Win32 API
'**************************************************************************************************
 Private Declare Function Blend Lib "gdi32.dll" Alias "GdiAlphaBlend" (ByVal desthDC As Long, _
     ByVal destX As Long, ByVal destY As Long, ByVal destW As Long, ByVal destH As Long, _
     ByVal srchDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcW As Long, _
     ByVal srcH As Long, ByVal BLENDFUNCT As Long) As Long
'Private Declare Function Blend Lib "msimg32" Alias "AlphaBlend" (ByVal desthDC As Long, _
'     ByVal destX As Long, ByVal destY As Long, ByVal destW As Long, ByVal destH As Long,
'     ByVal srchDC As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcW As Long,
'     ByVal srcH As Long, ByVal BLENDFUNCT As Long) As Long

'**************************************************************************************************
' KnightRider Module-Scoped Variables
'**************************************************************************************************
Implements WinSubHook2.iTimer
Private m_Tmr As cTimer

'**************************************************************************************************
' KnightRider Events
'**************************************************************************************************
Public Event TripComplete()

'**************************************************************************************************
' KnightRider Events
'**************************************************************************************************
Private m_Effect As eEffect
Private m_Enabled As Boolean
Private m_Height As Long
Private m_Position As Long
Private m_Speed As Long
Private m_Tail As Long
Private m_Width As Long

'**************************************************************************************************
' KnightRider Properties
'**************************************************************************************************
Public Property Get BackColor() As OLE_COLOR
     BackColor = picBlend.BackColor
End Property ' Get BackColor

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
     picBlend.BackColor = New_BackColor
     UserControl.BackColor = New_BackColor
     PropertyChanged "BackColor"
End Property ' Let BackColor

Public Property Get Effect() As eEffect
Attribute Effect.VB_Description = "Sets/returns the effect."
     Effect = m_Effect
End Property ' Get Effect

Public Property Let Effect(ByVal New_Effect As eEffect)
          Select Case New_Effect
               Case 0
                    m_Position = 0
                    m_Speed = Abs(m_Speed)
               Case 1
                    m_Position = ScaleWidth
                    m_Speed = -Abs(m_Speed)
               Case 2
                    m_Position = ScaleWidth / 2
               Case Else
                    Err.Raise 381, Me, "Invalid Effect"
                    Exit Property
          End Select
          m_Effect = New_Effect
          PropertyChanged "Effect"
End Property ' Let Effect

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets/returns whether the control is enabled."
Attribute Enabled.VB_UserMemId = 0
Attribute Enabled.VB_MemberFlags = "200"
     Enabled = m_Enabled
End Property ' Get Enabled

Public Property Let Enabled(ByVal New_Enabled As Boolean)
     If New_Enabled Then
          m_Tmr.TmrStart Me, 30
     Else
          m_Tmr.TmrStop
          Refresh
     End If
     m_Enabled = New_Enabled
     PropertyChanged "Enabled"
End Property ' Let Enabled

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
     ForeColor = picBlend.ForeColor
End Property ' Get ForeColor

Public Property Let ForeColor(ByVal nuForeColor As OLE_COLOR)
     picBlend.ForeColor() = nuForeColor
     PropertyChanged "ForeColor"
End Property ' Let ForeColor

Public Property Get Speed() As Long
Attribute Speed.VB_Description = "Sets/returns the speed. Usable values are 1 thru 10."
    Speed = Abs(m_Speed)
End Property ' Get Speed

Public Property Let Speed(ByVal New_Speed As Long)
     Select Case New_Speed
          Case Is <= 0
               Err.Raise 382, Me, "Invalid Speed"
          Case Else
               If m_Speed = 0 Then m_Speed = 1
               m_Speed = New_Speed * Sgn(m_Speed)
               PropertyChanged "Speed"
     End Select
End Property ' Let Speed

Public Property Get Tail() As Long
Attribute Tail.VB_Description = "Sets/returns the tail length."
     Tail = 31 - m_Tail \ 65536
End Property ' Get Tail

Public Property Let Tail(ByVal New_Tail As Long)
     m_Tail = (31 - (New_Tail And 31)) * 65536
     PropertyChanged "Tail"
End Property ' Tail

'**************************************************************************************************
' KnightRider Implemented TimerProc
'**************************************************************************************************
Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
     Blend hDC, 0, 0, m_Width, m_Height, picBlend.hDC, 0, 0, m_Width, m_Height, m_Tail
     Line (m_Position, 0)-(m_Position + m_Speed - 1, m_Height - 1), picBlend.ForeColor, BF
     m_Position = m_Position + m_Speed
     Select Case m_Effect
          Case 0
               If m_Position > m_Width Then m_Position = False
          Case 1
               If m_Position < 0 Then m_Position = m_Width
          Case 2
               If m_Position < 0 Or m_Position > m_Width Then m_Speed = -m_Speed
               If m_Position <= 0 Then RaiseEvent TripComplete
    End Select
End Sub ' iTimer_Proc
'**************************************************************************************************
' KnightRider Intrinsic Methods
'**************************************************************************************************
Private Sub UserControl_Initialize()
     ' Create timer object
     Set m_Tmr = New cTimer
End Sub ' UserControl_Initialize

Private Sub UserControl_InitProperties()
     Effect = 2
     Speed = 2
     Tail = 15
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     With PropBag
          BackColor = .ReadProperty("BackColor", vbBlack)
          picBlend.ForeColor = .ReadProperty("ForeColor", vbGreen)
          Enabled = .ReadProperty("Enabled", False)
          Effect = .ReadProperty("Effect", 2)
          Speed = .ReadProperty("Speed", 2)
          Tail = .ReadProperty("Tail", 15)
     End With
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     picBlend.Move 0, 0, Width, Height
     m_Width = ScaleWidth
     m_Height = ScaleHeight
     'this is here to avoid an error while effect is unknown
     If m_Effect Then Effect = m_Effect
End Sub ' UserControl_Resize

Private Sub UserControl_Terminate()
     ' Destroy our timer object
     m_Tmr.TmrStop
     Set m_Tmr = Nothing
End Sub ' UserControl_Terminate

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     With PropBag
          .WriteProperty "BackColor", picBlend.BackColor, vbBlack
          .WriteProperty "ForeColor", picBlend.ForeColor, vbGreen
          .WriteProperty "Enabled", m_Enabled, False
          .WriteProperty "Effect", m_Effect, 2
          .WriteProperty "Speed", Abs(m_Speed), 2
          .WriteProperty "Tail", Tail, 15
     End With
End Sub ' UserControl_WriteProperties

