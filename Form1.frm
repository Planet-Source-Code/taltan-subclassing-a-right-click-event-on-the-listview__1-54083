VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub classing by Taltan"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView MyListView 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Header"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "More header"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuPOP 
      Caption         =   "POPUP"
      Visible         =   0   'False
      Begin VB.Menu hey 
         Caption         =   "Hey"
      End
      Begin VB.Menu hello 
         Caption         =   "Hello"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Get the child of MyListView
EnumListVewChild MyListView.hWnd

'Change the new classing of the control:
lngOldProc = SetWindowLong(ListViewHeader_hWnd, GWL_WNDPROC, AddressOf SubClass)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'And of course, lets clean up after us :p
SetWindowLong ListViewHeader_hWnd, GWL_WNDPROC, lngOldProc
End Sub
