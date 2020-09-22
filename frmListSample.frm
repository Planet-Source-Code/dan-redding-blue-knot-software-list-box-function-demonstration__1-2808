VERSION 5.00
Begin VB.Form fmrListSample 
   Caption         =   "List Box Project Sample"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   Icon            =   "frmListSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   315
      Left            =   2340
      TabIndex        =   11
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Find ""&Begins With Item 1"""
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   2640
      Width           =   2235
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3360
      TabIndex        =   12
      Top             =   4080
      Width           =   1275
   End
   Begin VB.CommandButton cmdFindExact 
      Caption         =   "Find Item 1 &Exactly"
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   2235
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change Item 1 to Item 2"
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   2235
   End
   Begin VB.TextBox txtItem2 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   900
      Width           =   2235
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove Selected Item"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Item 1 To List"
      Default         =   -1  'True
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   1320
      Width           =   2235
   End
   Begin VB.TextBox txtItem1 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   300
      Width           =   2235
   End
   Begin VB.ListBox lstSample 
      Height          =   3960
      ItemData        =   "frmListSample.frx":0442
      Left            =   120
      List            =   "frmListSample.frx":0479
      TabIndex        =   9
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label lblResult 
      Height          =   855
      Left            =   2400
      TabIndex        =   8
      Top             =   3120
      Width           =   2235
   End
   Begin VB.Label lblSecond 
      Caption         =   "Item &2"
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label lblFirst 
      Caption         =   "Item &1"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "fmrListSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngLstIH As Long, lngLstWid As Long, lngLstMax As Long, lngScrollW As Long
    'These are form level because they are initialized once in Form_Load but used
    'in lstSample_MouseOver - this function is called repeatedly while the mouse is over
    'the listbox, and you don't want to recalculate them several times per second.

Private Sub cmdAdd_Click()
    If txtItem1.Text = "" Then Exit Sub 'Nothing to add!
    'Add item
    lstSample.AddItem txtItem1.Text
    'there is now definately something to clear
    cmdClear.Enabled = True
    'Select text in txtItem1 to make it easier to type next item over this one
    txtItem1.SelStart = 0
    txtItem1.SelLength = Len(txtItem1.Text)
End Sub

Private Sub cmdBegin_Click()
Dim lngPos As Long
Static lngLast As Long 'Doesn't reset automatically each time

    'is this the first time we've looked?  We can tell by checking the
    'caption we changed after the first time.
    If Left$(cmdBegin.Caption, 9) = "Find Next" Then
        lngPos = FindNextBeginsInList(lstSample, txtItem1.Text, lngLast)
    Else
        lngLast = -1 'reset in case we've searched for something else
        lngPos = FindBeginsWithInList(lstSample, txtItem1.Text)
    End If
    
        
    If lngPos >= 0 Then 'Found!
        If lngPos <= lngLast Then 'the search wraps around when
                            'it reaches the end
            lblResult = "The Entire List has been searched"
            cmdBegin.Enabled = False 'Turn off button till item 1 changed
        Else
            lblResult = lstSample.List(lngPos) & " was found in position" & Str$(lngPos)
        End If
        'change caption.  The Chr$(34)'s are " marks
        cmdBegin.Caption = "Find Next " & Chr$(34) & "&Begins With..." & Chr$(34)
    Else
        'not found at all!
        lblResult = "Nothing beginning with " & txtItem1.Text & " is in the list"
        cmdBegin.Enabled = False 'Turn it off till item 1 is changed
    End If
    lngLast = lngPos 'know where to start looking next time
End Sub

Private Sub cmdChange_Click()
    'function returns true if successful change was made
    If ChangeInList(lstSample, txtItem1.Text, txtItem2.Text) Then
        lblResult = "Item Changed"
    Else
        lblResult = txtItem1.Text & " is not in the list"
    End If
End Sub

Private Sub cmdClear_Click()
    'clear the list
    lstSample.Clear
    'there is now nothing to clear or remove!
    cmdClear.Enabled = False
    cmdRemove.Enabled = False
End Sub

Private Sub cmdExit_Click()
    'Quit
    Unload Me
End Sub

Private Sub cmdFindExact_Click()
Dim lngPos As Long
    lngPos = FindExactInList(lstSample, txtItem1.Text)
    'function returns index position or -1 if not found
    If lngPos >= 0 Then
        lblResult = txtItem1.Text & " is in position" & Str$(lngPos)
        lstSample.ListIndex = lngPos
    Else
        lblResult = txtItem1.Text & " is not in the list"
    End If
End Sub

Private Sub cmdRemove_Click()
    'removes selected item
    lstSample.RemoveItem lstSample.ListIndex
    'nothing is now selected --> nothing to remove!
    cmdRemove.Enabled = False
    'is there anything left in the list to clear?
    If lstSample.ListCount = 0 Then cmdClear.Enabled = False
End Sub

Private Sub Form_Load()
Dim lReturn As Long, recLst As RECT
    'Height in pixels of an item on the list
    lngLstIH = SendMessageByNum(lstSample.hwnd, LB_GETITEMHEIGHT, 0, 0)
    'get dimensions of listbox in pixels
    'lReturn is a dummy var.  The real data goes into recLst
    lReturn = GetClientRect(lstSample.hwnd, recLst)
    'Number of lines on the list
    lngLstMax = (recLst.Bottom - recLst.Top) \ lngLstIH
    'Width of inside of listbox (not including borders) in pixels
    lngLstWid = (recLst.Right - recLst.Left)
    'Width of a vertical ScrollBar in pixels
    lngScrollW = GetSystemMetrics(SM_CYVSCROLL)
    
    'If the list box resizes with your form, put these calcs in the listbox's Resize
    'event instead
End Sub

Private Sub lstSample_Click()
    'Something has been selected, it can be removed
    cmdRemove.Enabled = True
End Sub

Private Sub lstSample_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim idx As Integer, strItem As String
    'calculate index position by diving y (twips) by the height of a list item
    'by the height of a single item (converted to twips) and adding it to
    'TopIndex - the index of the value at the top of the list
    idx = lstSample.TopIndex + (Y \ (lngLstIH * Screen.TwipsPerPixelY))
    
    If idx > lstSample.ListCount - 1 Then Exit Sub 'pointing at blank space under items!
    
    strItem = lstSample.List(idx)
    'listboxes don't have TextWidth methods, but forms do.
    'make sure Font Properties are same for Form as for listbox
    'if they _need_ to be different, you can use the TextWidth method of a picturebox
    '(which can be hidden if you don't need it otherwise)
    
    'the following statement tests the width of the text (in Twips)
    'against the width of the inside of the listbox
    'The IIF statement returns the width of a scroll bar if one is showing
    'by texting to see if the listcount is greater than the max lines showing in
    'the listbox (or returns 0 if it is not).  That value is subtracted from the
    'width to give the 'visible' width.  That value is then converted to Twips.
    
    If Me.TextWidth(strItem) >= (lngLstWid - _
        IIf(lstSample.ListCount > lngLstMax, lngScrollW, 0)) _
        * Screen.TwipsPerPixelY Then
        'If true, then item pointed at is too wide to see
        lstSample.ToolTipText = strItem
    Else
        lstSample.ToolTipText = "" 'Turn it off if it was on!
    End If
End Sub

Private Sub txtItem1_Change()
    'New Item 1, start over with cmdBegin
    cmdBegin.Caption = "Find " & Chr$(34) & "&Begins With Item 1" & Chr$(34)
    cmdBegin.Enabled = True
End Sub
