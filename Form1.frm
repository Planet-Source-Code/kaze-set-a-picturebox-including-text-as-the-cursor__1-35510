VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   2820
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   2820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "reset cursor"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   1995
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "0987/2"
      Top             =   180
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "set box image as cursor"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "move text to PicBox"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawMode        =   9  'Not Mask Pen
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2220
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   180
      Width           =   510
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2160
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16777215
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'example by: Chris LaJoie
'if the text is too long, all of it won't show up, and
'I don't recommend resizing the picturebox.  Cursors
'are designed to be 32x32 pixels and if you attempt to
'make it smaller/larger than that, windows will squish
'it or stretch it to make it 32x32.  The best thing to
'do is just leave the picturebox at 32x32 pixels, that
'you know you're not going over/under.  however, you
'can change the font size, and print text on 2 lines of
'the picturebox if you want.

Private Sub Command1_Click()
    Picture1.Cls 'clearing the picture in case you added something before
    Picture1.Print Text1.Text 'printing the text
    
    imgList.ListImages.Clear 'clearing the image list control
    imgList.ListImages.Add , , Picture1.Image 'since an index is not specified, the image will be added to index 1
End Sub

Private Sub Command2_Click()
'the reason I use an imagelist is because it has an ExtractIcon
'feature which allows you to convert the picture to an icon,
'allowing you to set it as the cursor.
    Set Me.MouseIcon = imgList.ListImages.Item(1).ExtractIcon 'extracting the icon, and seting it as the custom mouseicon
    Me.MousePointer = vbCustom 'allowing the form to display a custom mouse pointer

End Sub

Private Sub Command3_Click()
    Me.MousePointer = vbDefault
End Sub
