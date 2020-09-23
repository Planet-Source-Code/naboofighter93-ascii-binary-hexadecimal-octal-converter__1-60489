VERSION 5.00
Begin VB.Form Wait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1065
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Wait.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Calculating, Please wait"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1680
      End
   End
End
Attribute VB_Name = "Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
