VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTooltip3 
      Interval        =   100
      Left            =   4935
      Top             =   6195
   End
   Begin VB.Timer tmrTooltip2 
      Interval        =   100
      Left            =   4410
      Top             =   6195
   End
   Begin VB.Timer tmrTooltip1 
      Interval        =   100
      Left            =   3885
      Top             =   6195
   End
   Begin VB.Frame frameTTip2 
      BackColor       =   &H80000017&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   200
      Left            =   315
      TabIndex        =   17
      Top             =   6510
      Width           =   200
      Begin VB.Shape Shape2 
         BackColor       =   &H80000017&
         BorderColor     =   &H00000000&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   140
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   140
      End
   End
   Begin VB.Frame frameTTip1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   100
      Left            =   105
      TabIndex        =   16
      Top             =   6510
      Width           =   100
   End
   Begin VB.Frame frameTTip3 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   225
      Left            =   650
      TabIndex        =   14
      Top             =   6500
      Width           =   3060
      Begin VB.Label LabelTooltip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabelTooltip"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3060
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Delete a node: Enter the key!"
      Height          =   1065
      Left            =   3990
      TabIndex        =   11
      Top             =   3255
      Width           =   2535
      Begin VB.TextBox txtDelete 
         Height          =   285
         Left            =   315
         TabIndex        =   13
         Text            =   "Enter node key here"
         Top             =   680
         Width           =   1800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   330
         Left            =   420
         TabIndex        =   12
         ToolTipText     =   "Click to delete a node"
         Top             =   280
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add a child node"
      Height          =   2010
      Left            =   3990
      TabIndex        =   5
      Top             =   1155
      Width           =   2535
      Begin VB.TextBox txtNodeKey 
         Height          =   285
         Left            =   945
         TabIndex        =   18
         ToolTipText     =   "Key is needed to identify a node!"
         Top             =   735
         Width           =   1485
      End
      Begin VB.TextBox txtNodeTo 
         Height          =   285
         Left            =   945
         TabIndex        =   8
         Text            =   "Offline"
         ToolTipText     =   "The parent node you want to add the node"
         Top             =   1155
         Width           =   1485
      End
      Begin VB.TextBox txtNodeText 
         Height          =   285
         Left            =   945
         TabIndex        =   7
         ToolTipText     =   "Text appears in the tree!"
         Top             =   315
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Node"
         Height          =   330
         Left            =   525
         TabIndex        =   6
         ToolTipText     =   "Click to add a node"
         Top             =   1575
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "Node Key:"
         Height          =   225
         Left            =   105
         TabIndex        =   19
         Top             =   735
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "To:"
         Height          =   225
         Left            =   525
         TabIndex        =   10
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Node Text:"
         Height          =   225
         Left            =   105
         TabIndex        =   9
         Top             =   315
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1485
      Left            =   3990
      TabIndex        =   1
      Top             =   4410
      Width           =   2535
      Begin VB.TextBox txtto 
         Height          =   285
         Left            =   525
         TabIndex        =   4
         Text            =   "To"
         ToolTipText     =   "To"
         Top             =   630
         Width           =   1380
      End
      Begin VB.TextBox txtchild 
         Height          =   285
         Left            =   525
         TabIndex        =   3
         Text            =   "Child"
         ToolTipText     =   "Child"
         Top             =   260
         Width           =   1380
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "Move"
         Height          =   330
         Left            =   525
         TabIndex        =   2
         ToolTipText     =   "Click to move node from one parent to another"
         Top             =   1020
         Width           =   1380
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5775
      Top             =   6090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":00D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":01A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0270
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0638
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4635
      Left            =   210
      TabIndex        =   0
      Top             =   1260
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8176
      _Version        =   393217
      Style           =   5
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label lblAbout 
      Caption         =   $"Form1.frx":0710
      Height          =   960
      Left            =   105
      TabIndex        =   20
      Top             =   105
      Width           =   6420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' ===============================================
' ===============================================
' ===============================================
'This is a demonstration of how to manage a
'treeview, i.e. add, delete and
'move nodes of the treeview. But more than
'that it also shows how can we make coooool
'tooltips. People who are tired to see
'those traditional tooltips can now customise
'the look of tooltips as shown in this example.
'Just move your mouse over the nodes of the treeview below.

'Cheers!
'Irfan Ullah Khan
'skilljaan@hotmail.com
'"The world needs Khilafah!"

' NOTE: IF YOU COPY AND PASTE THIS CODE INTO YOUR
'       APPLICATIONS, IT WILL WORK BUT PERHAPS YOU
'       WILL HAVE TO MODIFY ONE PART OF THIS CODE; (THAT
'       PART IS SHOWN BELOW)
' ===============================================
' ===============================================
' ===============================================


Private Sub cmdMove_Click()
Set TreeView1.Nodes(txtchild.Text).Parent = TreeView1.Nodes(txtto.Text)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        frameTTip3.Visible = False
        frameTTip2.Visible = False
        frameTTip1.Visible = False
End Sub

Private Sub LabelTooltip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frameTTip3.Visible = False
End Sub



Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        frameTTip3.Visible = False
        frameTTip2.Visible = False
        frameTTip1.Visible = False
End Sub

Private Sub tmrTooltip1_Timer()
frameTTip1.Visible = True
tmrTooltip1.Enabled = False
tmrTooltip2.Enabled = True
End Sub

Private Sub tmrTooltip2_Timer()
frameTTip2.Visible = True
tmrTooltip2.Enabled = False
tmrTooltip3.Enabled = True
End Sub

Private Sub tmrTooltip3_Timer()
frameTTip3.Visible = True
tmrTooltip3.Enabled = False
'Timer3.Enabled = True
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objNode As Node

    Set objNode = TreeView1.HitTest(X, Y)

    If Not objNode Is Nothing Then
' NOTE: IF YOU COPY AND PASTE THIS APPLICATION'S CODE
'       INTO YOUR APPLICATIONS, IT WILL WORK BUT
'       PERHAPS YOU WILL HAVE TO MODIFY THE FOLLOWING
'       PART (ACCORDING TO THE PLACEMENT OF THE
'       TREEVIEW IN THE FORM):
            frameTTip1.Move X + 250, Y + 1180
            frameTTip2.Move X + 280, Y + 1000
            frameTTip3.Move X + 400, Y + 730
        ' NOTE: All these values are in twips!
        '       If the ScaleMode property of your
        '       form is set to "3 - Pixel", then you
        '       will have to modify the above
        '       as follows (to change values to pixels):
        '   frameTTip1.Move (X + 250) / Screen.TwipsPerPixelX, (Y + 1180) / Screen.TwipsPerPixelY
        '   frameTTip2.Move (X + 280) / Screen.TwipsPerPixelX, (Y + 1000) / Screen.TwipsPerPixelY
        '   frameTTip3.Move (X + 400) / Screen.TwipsPerPixelX, (Y + 730) / Screen.TwipsPerPixelY
        
        tmrTooltip1.Enabled = True ' Enable the timer!!
        ' Now set the text of label to the key of node (I just added a space before it so that it does not look like touching the edge)!
        LabelTooltip.Caption = " " & objNode.Key
        ' If label text is Online/Offline i.e. if cursor is on those parent nodes dont show the tooltip!
        If LabelTooltip.Caption = " Online" Or LabelTooltip.Caption = " Offline" Then
                frameTTip3.Visible = False
                frameTTip2.Visible = False
                frameTTip1.Visible = False
        End If
        
        'Set the width size of the label so that it adjusts according to the length of the key values of the nodes!
        Select Case Len(LabelTooltip.Caption) '
        Case Is < 10:
        frameTTip3.Width = 755
        LabelTooltip.Width = 755
        Case 10:
        frameTTip3.Width = 835
        LabelTooltip.Width = 835
        Case 11:
        frameTTip3.Width = 925
        LabelTooltip.Width = 925
        Case 12:
        frameTTip3.Width = 1000
        LabelTooltip.Width = 1000
        Case 13:
        frameTTip3.Width = 1085
        LabelTooltip.Width = 1085
        Case 14:
        frameTTip3.Width = 1170
        LabelTooltip.Width = 1170
        Case 15:
        frameTTip3.Width = 1245
        LabelTooltip.Width = 1245
        Case 16:
        frameTTip3.Width = 1330
        LabelTooltip.Width = 1330
        Case 17:
        frameTTip3.Width = 1415
        LabelTooltip.Width = 1415
        Case 18:
        frameTTip3.Width = 1500
        LabelTooltip.Width = 1500
        Case 19:
        frameTTip3.Width = 1585
        LabelTooltip.Width = 1585
        Case 20:
        frameTTip3.Width = 1670
        LabelTooltip.Width = 1670
        Case 21:
        frameTTip3.Width = 1840
        LabelTooltip.Width = 1840
        Case 22:
        frameTTip3.Width = 1925
        LabelTooltip.Width = 1925
        Case 23:
        frameTTip3.Width = 2010
        LabelTooltip.Width = 2010
        Case 24:
        frameTTip3.Width = 2095
        LabelTooltip.Width = 2095
        Case 25:
        frameTTip3.Width = 2180
        LabelTooltip.Width = 2180
        Case 26:
        frameTTip3.Width = 2265
        LabelTooltip.Width = 2265
        Case 27:
        frameTTip3.Width = 2350
        LabelTooltip.Width = 2350
        Case 28:
        frameTTip3.Width = 2435
        LabelTooltip.Width = 2435
        Case 29:
        frameTTip3.Width = 2520
        LabelTooltip.Width = 2520
        Case 30:
        frameTTip3.Width = 2600
        LabelTooltip.Width = 2600
        
        Case Is > 30:
        frameTTip3.Width = 2700
        LabelTooltip.Width = 2700
        End Select
               
    End If
    
    'If the cursor is not above any node, dont show the tooltip!
    If objNode Is Nothing Then
        frameTTip3.Visible = False
        frameTTip2.Visible = False
        frameTTip1.Visible = False
    End If
End Sub

Private Sub Command2_Click()
On Error GoTo errrr:
TreeView1.Nodes.Remove txtDelete.Text
Exit Sub
errrr:
MsgBox "No such node!"
End Sub

Private Sub Form_Load()

Me.Width = 6800
Me.Height = 6500

' Add these nodes when the form loads!
TreeView1.Nodes.Add , , "Online", "Online", 2
TreeView1.Nodes.Add , , "Offline", "Offline", 1
TreeView1.Nodes.Add "Online", tvwChild, "lailahaillallah@hotmail.com", "La ilaha illallah", 2
TreeView1.Nodes.Add "Online", tvwChild, "allah0akbar@hotmail.com", "Allah-0-Akbar", 2
TreeView1.Nodes.Add "Online", tvwChild, "skilljaan@hotmail.com", "Erfan", 2
TreeView1.Nodes.Add "Online", tvwChild, "usamaalam@hotmail.com", "Usama", 3
TreeView1.Nodes.Add "Online", tvwChild, "Yuwiiieee@hotmail.com", "Yuweeeee", 5
TreeView1.Nodes.Add "Online", tvwChild, "busharraf@hotmail.com", "Busharraf", 4

TreeView1.Nodes.Add "Offline", tvwChild, "yessssss@hotmail.com", "Custom Tooltips", 1
TreeView1.Nodes.Add "Offline", tvwChild, "cooltooltips@hotmail.com", "Coooooool", 1
TreeView1.Nodes.Add "Offline", tvwChild, "khilafah@khilafah.com", "The world needs Khilafah", 1


TreeView1.Nodes("Offline").Expanded = True
TreeView1.Nodes("Online").Expanded = True

ControlToEllipse frameTTip1, 3, 3  'rounded tooltips
ControlToEllipse frameTTip2, 10, 10 'rounded tooltips

End Sub

Private Sub Command3_Click()
TreeView1.Nodes.Add txtNodeTo.Text, tvwChild, txtNode.Text, txtNode.Text, 2
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
        If TreeView1.SelectedItem.Key = "Offline" Or TreeView1.SelectedItem.Key = "Online" Then
        Exit Sub
        Else
        txtto.Text = TreeView1.SelectedItem.Key
        End If

End Sub

'========= Code to make the round shapes ==========
'=============== rounded shapes ===================
Private Sub ControlToEllipse(aControl As Variant, Width As Integer, Height As Integer)
    Dim l As Long
    l = CreateEllipticRgn(0, 0, Width, Height)
    SetWindowRgn aControl.hwnd, l, True
End Sub

Private Sub txtDelete_GotFocus()
txtDelete.Text = ""
End Sub
