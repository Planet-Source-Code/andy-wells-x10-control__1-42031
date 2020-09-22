VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "For more information, read the readme file."
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   2955
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "X10 Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AIM:  Trumpet Wellsy"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      Caption         =   "Email:  awells@comnetcom.net"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2205
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Created by:  Andy Wells"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
