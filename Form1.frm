VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "vbXML Compatible Example"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Full XML"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Node Lists"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Color Codes"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Encryption / Decryption"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtEnc"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdEnc"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdDec"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cmdDec 
         Caption         =   "Decrypt Data"
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdEnc 
         Caption         =   "Encrypt Data"
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtEnc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   6735
      End
      Begin VB.Frame Frame3 
         Caption         =   "RGB Color Codes"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   6735
         Begin VB.TextBox txtText 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox txtList 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            TabIndex        =   10
            Top             =   480
            Width           =   3255
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save Color Codes"
            Height          =   255
            Left            =   2400
            TabIndex        =   9
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "TextBox Code: (Fore - Back)"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label4 
            Caption         =   "ListView Code: (Fore - Back)"
            Height          =   255
            Left            =   3360
            TabIndex        =   12
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Caption         =   "Parent Nodes Under <main>"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   6735
         Begin MSComctlLib.ListView lvColors 
            Height          =   4095
            Left            =   3840
            TabIndex        =   4
            Top             =   480
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   7223
            View            =   2
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvMain 
            Height          =   4095
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   7223
            View            =   2
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Nodes under <colors>"
            Height          =   255
            Left            =   3840
            TabIndex        =   7
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Nodes under <main>"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Full XML Code"
         Height          =   4815
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton cmdSaveXML 
            Appearance      =   0  'Flat
            Caption         =   "Save XML File"
            Height          =   255
            Left            =   4680
            TabIndex        =   14
            Top             =   4440
            Width           =   1935
         End
         Begin VB.TextBox txtXML 
            Appearance      =   0  'Flat
            Height          =   4215
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   240
            Width           =   6495
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Node Data: (/main/data)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDec_Click()
XML.DecryptNode ("/main/data")

txtEnc.Text = XML.ReadNode("/main/data")

XML.Save ("vbXML.xml")
End Sub

Private Sub cmdEnc_Click()
XML.EncryptNode ("/main/data")

txtEnc.Text = XML.ReadNode("/main/data")

XML.Save ("vbXML.xml")
End Sub

Private Sub cmdSave_Click()
' Write TextBox values
XML.WriteNode "/main/colors/tbox/fore", Left(txtText, InStr(txtText, " - ") - 1)
XML.WriteNode "/main/colors/tbox/back", Mid(txtText, InStr(txtText, " - ") + 3)

' Write ListView values
XML.WriteNode "/main/colors/lview/fore", Left(txtList, InStr(txtList, " - ") - 1)
XML.WriteNode "/main/colors/lview/back", Mid(txtList, InStr(txtList, " - ") + 3)

' Reload the XML File after saving
XML.Save App.Path & "\vbxml.xml"
InitXML
End Sub

Private Sub cmdSaveXML_Click()
XML.OpenXML txtXML.Text, oxString
XML.Save App.Path & "\vbxml.xml"

InitXML
End Sub

Private Sub Form_Load()
'---------------------------------------------------
' vbXML uses the MSXML feature XPath.  Here is a quick
' example of how I would query a node (used in
' ReadNode, ReadNodeXML, and WriteNode, and many others):
' To access a node:
'   "/parent1/childofparent1/childofparent2"
' Here is a quick explination:
' I am using test.xml for this example.  For ease of
' this example I have pasted the contents here:
'
'  <test>
'     <text>
'         <hello>Hello</hello>
'         <bye>Goodbye</bye>
'     </text>
'  </test>
'
' To access the <hello> node, you would use this
' query:
'   "/test/text/hello"
'
' To access the <bye> node, you would use this query:
'   "/test/text/bye"
'
' In the last example (where we queried the <bye> node)
' "/test/" is parent1, "/text/" is childofparent1,
' and "/bye" is childofparent2
'
' Take note: you can have multiple child nodes
' (I dont know the exact count)
'---------------------------------------------------

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' This example is an XML editor.  It is meant to run
' with the vbXML class wrapper for Visual Basic
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

SSTab1.Tab = 0

' See InitXML (In the basMain module) for the XML loading
InitXML
End Sub

Private Sub lvColors_Click()
SSTab1.SetFocus
End Sub

Private Sub lvMain_Click()
SSTab1.SetFocus
End Sub
