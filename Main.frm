VERSION 5.00
Begin VB.Form Main 
   Caption         =   "AIM Password Retriever"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Get Passwords"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   3735
   End
   Begin VB.ListBox List2 
      Height          =   2010
      ItemData        =   "Main.frx":0000
      Left            =   2040
      List            =   "Main.frx":0002
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "Main.frx":0004
      Left            =   120
      List            =   "Main.frx":0006
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Password"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Screennames"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########################################
'# Info                                    #
'#                                         #
'# Creator: ChiChis                        #
'# Date: Jul. 25, 2000                     #
'# Made with help/example from:            #
'#   Algorithm description-                #
'#     http://www.tlsecurity.net           #
'#   Registry class module-                #
'#     Riaan Aspeling                      #
'###########################################
'# Description                             #
'#                                         #
'# This code encapsulates everything you   #
'# need to extract AIM screennames and     #
'# passwords.  It has the code to get all  #
'# the registry keys and then decrypt the  #
'# password the very long algorithm.       #
'###########################################
'# How To Use                              #
'#                                         #
'# GetAIMs SNListBox, PWListBox            #
'# It's simple to use, and can easily be   #
'# altered to display the results in a     #
'# textbox.                                #
'###########################################

Private Sub Command1_Click()
    GetAIMs List1, List2
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.ListIndex = List2.ListIndex
End Sub
