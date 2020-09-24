VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7335
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "Run Test"
      Height          =   495
      Left            =   2677
      TabIndex        =   1
      Top             =   3975
      Width           =   1980
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   360
      TabIndex        =   0
      Top             =   450
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Table properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   135
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Option Base 1 'uncomment to see the difference

Private TestTable1(1 To 3, 4, -5 To 5, -6 To -5, 2000) As String
Private TestTable2() As String * 8

Private Sub Command1_Click()

    List1.Clear
    List1.AddItem "TestTable1 (dimmed)"
    List1.AddItem "  IsDimmed    " & IsDimmed(hTable(TestTable1))
    DisplayTableDetails GetArrayDescriptor(hTable(TestTable1))

    List1.AddItem ""
    List1.AddItem "TestTable2 (un-dimmed)"
    List1.AddItem "  IsDimmed    " & IsDimmed(hTable(TestTable2))
    DisplayTableDetails GetArrayDescriptor(hTable(TestTable2))

    List1.AddItem ""
    ReDim TestTable2(1 To 13, 8, -5 To 5, -6 To -6, 100)
    List1.AddItem "TestTable2 (redimmed)"
    List1.AddItem "  IsDimmed    " & IsDimmed(hTable(TestTable2))
    DisplayTableDetails GetArrayDescriptor(hTable(TestTable2))

    List1.AddItem ""
    ReDim TestTable2(1 To 1000)
    List1.AddItem "TestTable2 (redimmed again)"
    List1.AddItem "  IsDimmed    " & IsDimmed(hTable(TestTable2))
    DisplayTableDetails GetArrayDescriptor(hTable(TestTable2))

    List1.AddItem ""
    Erase TestTable2
    List1.AddItem "TestTable2 (erased)"
    List1.AddItem "  IsDimmed    " & IsDimmed(hTable(TestTable2))
    DisplayTableDetails GetArrayDescriptor(hTable(TestTable2))

End Sub

Private Sub DisplayTableDetails(ArrayDescriptor As SAFEARRAYDESCRIPTOR)

  Dim i As Long

    With ArrayDescriptor
        List1.AddItem "  Fixed       " & IIf(.Features And FADF_FIXEDSIZE, "Yes", "No")
        List1.AddItem "  Features    " & "x" & Hex$(.Features And Not FADF_RESERVED)
        List1.AddItem "  ElementSize " & .ElementSize
        List1.AddItem "  Dimensions  " & .NumDims
        For i = 1 To .NumDims
            With .Bounds(i)
                List1.AddItem "    Dimension " & i & ": LBound " & .LBound & ", UBound " & .UBound & ", Elements " & .NumElements
            End With '.BOUNDS(I)
        Next i
    End With 'ArrayDescriptor

End Sub

':) Ulli's VB Code Formatter V2.17.3 (2004-Jul-23 11:12) 6 + 56 = 62 Lines
