VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "VBC 1: VB Basics"
   ClientHeight    =   6105
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraYourText 
      Caption         =   "&Enter Your text here:"
      Height          =   735
      Left            =   0
      TabIndex        =   25
      Top             =   3960
      Width           =   6015
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Text            =   "YourText"
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame fraTextnFonts 
      Caption         =   "Text && Fonts:"
      Height          =   3855
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   3255
      Begin VB.Frame fraTextColor 
         Caption         =   "Fore && Backgroundcolor:"
         Height          =   2535
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   3015
         Begin VB.PictureBox picColor 
            BackColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   2160
            ScaleHeight     =   915
            ScaleWidth      =   675
            TabIndex        =   22
            Top             =   960
            Width           =   735
         End
         Begin VB.HScrollBar hscRGB 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   480
            Max             =   255
            TabIndex        =   21
            Top             =   1680
            Width           =   1575
         End
         Begin VB.HScrollBar hscRGB 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   480
            Max             =   255
            TabIndex        =   19
            Top             =   1320
            Width           =   1575
         End
         Begin VB.HScrollBar hscRGB 
            Height          =   255
            Index           =   0
            LargeChange     =   10
            Left            =   480
            Max             =   255
            TabIndex        =   17
            Top             =   960
            Width           =   1575
         End
         Begin VB.Frame fraColortype 
            Caption         =   "&Colortype:"
            Height          =   615
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2775
            Begin VB.OptionButton optColor 
               Caption         =   "BackColor"
               Height          =   195
               Index           =   1
               Left            =   1440
               TabIndex        =   15
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optColor 
               Caption         =   "ForeColor"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.Label lblColor 
            Height          =   255
            Left            =   1200
            TabIndex        =   24
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label lblRGB 
            Caption         =   "R/G/B Color:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label lblB 
            Caption         =   "&B:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label lblG 
            Caption         =   "&G:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblR 
            Caption         =   "&R:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   255
         End
      End
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   720
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cboFonts 
         Height          =   315
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "cboFonts"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblFontSize 
         Caption         =   "&Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblFontName 
         Caption         =   "&Font:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraBoxesnButtons 
      Caption         =   "CheckBoxes && OptionButtons"
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CheckBox chkBorder 
         Caption         =   "Border"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Frame fraAlignment 
         Caption         =   "A&lignment:"
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton optAlignment 
            Caption         =   "Center"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optAlignment 
            Caption         =   "Right"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optAlignment 
            Caption         =   "Left"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.CheckBox chkAppearance 
         Caption         =   "3D &Appearance"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Value           =   1  'Aktiviert
         Width           =   1935
      End
      Begin VB.CheckBox chkTransparent 
         Caption         =   "&Transparent"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Line lin3D_2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   2400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line lin3D_1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   120
         X2              =   2400
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin VB.Label lblOutput 
      Caption         =   "Your Text"
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   4800
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------
'!!!!!!!!!!! USE IT !!!!!!!!!!!!!!!!!!
'It's the best for a good programmer.
'The more variables you have the more you can lose track
'of things.
'If you're working on a bigger project you can have hundreds
'of variables and it's not difficult to forget a letter
'while writing a name, just like ImageCountr instead of
'ImageCounter and your program won't work and you'll spend
'the rest of your life wondering why it wasn't doing
'what it had to do. With Option Explicit it won't happen
'because every variabe, which isn't declared raises an
'error.
'---------------------------------------

Dim TextColor(0 To 1, 0 To 2) As Integer
'this will hold our Foreground and Background-Color
'informations for the label
'The first index (0 to 1) stands for Fore(0)/Back(1)
'the second for the R(0), G(1), B(2) - values

Dim CurrentColor As Byte
'It holds the selected color (0-ForeColor, 1-BackColor)

Private Sub txtText_Change()
'Each Object in VB has its "main" property like
'Caption ("main" property of a Label) or
'Text    ("main" property of a TextBox) or
'the selected Item in a Combo or ListBox
'It means, you can assign a value to this property or
'get a value from this property using only the Name of
'the Object: just like:
lblOutput = txtText
'Same as:
'lblOutput.Caption = txtText.Text
End Sub

Private Sub cboFonts_Click()
lblOutput.FontName = cboFonts
'Assign the choosen Font to the label
End Sub

Private Sub cboFontSize_Click()
lblOutput.FontSize = cboFontSize
'Do the same with the FontSize
End Sub

Private Sub chkAppearance_Click()
 lblOutput.Appearance = chkAppearance
 'sets the Appearance of the label(0-2D, 1-3D)
 RestoreColors
 'and restore the colors
End Sub

Private Sub chkBorder_Click()
 lblOutput.BorderStyle = chkBorder.Value
 'set the BorderStyle (0-No Border, 1-Border)
End Sub

Private Sub chkTransparent_Click()

'set the Transparancy of the label:
'using the property BackStyle(0-Transparent, 1-Not Transparent)

If chkTransparent.Value = 1 Then
 lblOutput.BackStyle = 0
Else
 lblOutput.BackStyle = 1
End If
  
'and restore colors:
 RestoreColors
End Sub

Private Sub Form_Load()
Dim i As Integer
'Just ask the Object: Screen how many fonts we have on this
'computer and add them all to the ComboBox cboFonts:
For i = 0 To Screen.FontCount - 1
  cboFonts.AddItem Screen.Fonts(i)
Next
'Add some FontSizes to the ComboBox cboFontSize
For i = 6 To 11
 cboFontSize.AddItem i
Next i
For i = 12 To 72 Step 4
 cboFontSize.AddItem i
Next
For i = 80 To 104 Step 8
 cboFontSize.AddItem i
Next i

'The ListIndex - property of a ComboBox returns/sets
'the currently choosen item
'(NOTE: First Item = 0, Last Item = ListCount-1)
cboFontSize.ListIndex = 3
cboFonts.ListIndex = 5
'the property Text of a ComboBox returns its choosen item
'Same as:
'cboFonts.List(cboFonts.ListIndex) or only
'cboFonts   (because the selected item is its "main" property)

lblOutput.FontName = cboFonts.Text
lblOutput.FontSize = cboFontSize

'and create the colors:
RestoreColors
End Sub

Private Sub Form_Unload(Cancel As Integer)
'This sub is called each time an user wants to quit(Unload)
'the Form/Project
'Also if he presses the X-Button in the top right corner
Dim answer As Integer 'will hold the answer of the user

answer = MsgBox("Quit ?", vbQuestion + vbYesNo, "VBC 1")
'ask the user if he wants to quit?

If answer = vbNo Then Cancel = 1
'if no, set Cancel=1, and the form won't unload
End Sub

Private Sub hscRGB_Change(Index As Integer)
'This sub handles every change of the 3 horizontal ScrollBars
'First the Color-Value is saved in the TextColor variable
TextColor(CurrentColor, Index) = hscRGB(Index)

'Show the values to the user
lblColor = hscRGB(0) & "/" & hscRGB(1) & "/" & hscRGB(2)

'And now show the color to the user
'The function RGB mixes the Red, Green and Blue values
'and returns the color number
picColor.BackColor = RGB(hscRGB(0), hscRGB(1), hscRGB(2))

'check which color have to be changed:
If CurrentColor = 1 Then
 lblOutput.BackColor = picColor.BackColor
Else
 lblOutput.ForeColor = picColor.BackColor
End If

End Sub

Private Sub mnuQuit_Click()
'Call the Form_Unload sub
Unload Me
End Sub

Private Sub optAlignment_Click(Index As Integer)
lblOutput.Alignment = Index
'sets the Alignment of the Label (0-Left, 1-Right, 2-Center)
End Sub

Private Sub optColor_Click(Index As Integer)
'set the color mode: (0-ForeColor, 1-BackColor)
CurrentColor = Index
'The "main" property of a ScrollBar is Value:
hscRGB(0) = TextColor(CurrentColor, 0)
hscRGB(1) = TextColor(CurrentColor, 1)
hscRGB(2) = TextColor(CurrentColor, 2)
'Show the values to the user:
lblColor = hscRGB(0) & "/" & hscRGB(1) & "/" & hscRGB(2)
'And create the color
picColor.BackColor = RGB(hscRGB(0), hscRGB(1), hscRGB(2))
End Sub

Private Sub txtText_GotFocus()
'This sub is called when the user selects this object
'It's a cool sub for TextBoxes, because it makes
'the Box mark the text, so you don't have to delete
'the whole text using Backspace!
txtText.SelStart = 0
txtText.SelLength = Len(txtText)
End Sub

'This sub is needed because after changing of some
'properties of the label the Backcolor may be changed
'to white or grey
Public Sub RestoreColors()
lblOutput.BackColor = RGB(TextColor(1, 0), TextColor(1, 1), TextColor(1, 2))
lblOutput.ForeColor = RGB(TextColor(0, 0), TextColor(0, 1), TextColor(0, 2))
End Sub
