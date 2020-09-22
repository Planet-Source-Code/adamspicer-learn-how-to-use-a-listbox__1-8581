VERSION 5.00
Begin VB.Form frmLISTBOX 
   Caption         =   "ListBox Example"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox lstNames 
      Height          =   1425
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "txtName"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "# of Names"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblNum 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblNum"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   3840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Name to add..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmLISTBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    lstNames.AddItem (txtName.Text) 'adds the name in the textbox to the listbox
    txtName.Text = "" 'clears the name out of the textbox
    lblNum.Caption = lstNames.ListCount 'puts how many names are in listbox into the label that will display the info

End Sub

Private Sub cmdClear_Click()
    lstNames.Clear 'clears the listbox

End Sub

Private Sub cmdRemove_Click()
    Dim Indx As Integer 'holds in memory the index of the list
    
    Indx = lstNames.ListIndex 'gets the index
    If Indx >= 0 Then 'makes sure that something is selected
        lstNames.RemoveItem Indx  'removes the selected item
        lblNum.Caption = lstNames.ListCount 'updates the number of names caption
    Else
        MsgBox "You must select an item to select!", vbExclamation, "Error In Delete" 'tells u why it wont work
    End If
    
    cmdRemove.Enabled = False

End Sub

Private Sub Form_Load()
    txtName.Text = "" 'clears the Name textbox
    lstNames.Text = "" 'clears the names listbox
    lblNum.Caption = "" 'clears the number of names label
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    
End Sub

Private Sub lstNames_Click()
    cmdRemove.Enabled = (lstNames.ListIndex <> -1) 'only enables command button if the list has something selected
    
End Sub

Private Sub txtName_Change()
    cmdAdd.Enabled = (Len(txtName.Text) > 0) 'only enables the add button if there is text in textbox

End Sub
