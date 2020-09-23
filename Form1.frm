VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00664444&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#"
   ClientHeight    =   8415
   ClientLeft      =   5385
   ClientTop       =   2385
   ClientWidth     =   7470
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   7470
   Begin VB.CheckBox chkfullCRC 
      BackColor       =   &H00664444&
      Caption         =   "Full CRC check"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Verify 
      Caption         =   "Verify"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton NewED 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   7560
      Width           =   1335
   End
   Begin VB.HScrollBar fixnumprc 
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox fixprefix 
      Appearance      =   0  'Flat
      BackColor       =   &H003C1009&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Top             =   5640
      Width           =   2655
   End
   Begin VB.OptionButton mode 
      BackColor       =   &H00664444&
      Caption         =   "Decode"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   2
      Top             =   110
      Width           =   975
   End
   Begin VB.OptionButton mode 
      BackColor       =   &H00664444&
      Caption         =   "Encode"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Top             =   110
      Width           =   975
   End
   Begin VB.ListBox lstStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H003C1009&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1470
      ItemData        =   "Form1.frx":0ECA
      Left            =   1920
      List            =   "Form1.frx":0ECC
      TabIndex        =   5
      Top             =   6000
      Width           =   5415
   End
   Begin VB.TextBox fixsavedir 
      Appearance      =   0  'Flat
      BackColor       =   &H003C1009&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton StartED 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   7560
      Width           =   1335
   End
   Begin MSComctlLib.ListView list 
      Height          =   3960
      Left            =   135
      TabIndex        =   15
      Top             =   375
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   6985
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   3936265
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   8040
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Line Line4 
      X1              =   7320
      X2              =   7320
      Y1              =   5160
      Y2              =   6000
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   5160
      Y2              =   7470
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   120
      Y1              =   7455
      Y2              =   7455
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   120
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAPID REPAIR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   270
      TabIndex        =   29
      Top             =   6600
      Width           =   1530
   End
   Begin VB.Image RRLogo 
      Height          =   960
      Left            =   480
      Picture         =   "Form1.frx":0ECE
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   960
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   6720
      TabIndex        =   28
      Top             =   4440
      Width           =   465
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   4080
      TabIndex        =   27
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   6720
      TabIndex        =   26
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   4080
      TabIndex        =   25
      Top             =   4440
      Width           =   465
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   1200
      TabIndex        =   24
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   1200
      TabIndex        =   23
      Top             =   4440
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fixfiles"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   22
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Renamed"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5760
      TabIndex        =   21
      Top             =   4440
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Missing"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3120
      TabIndex        =   20
      Top             =   4800
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Corrupt"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5760
      TabIndex        =   19
      Top             =   4800
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blocksize"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3120
      TabIndex        =   18
      Top             =   4440
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sourcefiles"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   17
      Top             =   4440
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To start, drag and drop files in the window below"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3405
   End
   Begin VB.Label lblfixnumprcval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2400
      TabIndex        =   12
      Top             =   5280
      Width           =   90
   End
   Begin VB.Label lblfixprefix 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repair file prefix"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3000
      TabIndex        =   11
      Top             =   5640
      Width           =   1140
   End
   Begin VB.Label lblPrc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   7680
      Width           =   45
   End
   Begin VB.Label lblfixsavedir 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create repair files in"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3000
      TabIndex        =   8
      Top             =   5280
      Width           =   1410
   End
   Begin VB.Label lblfixnumprc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Repair files to create"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   5280
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H003C1009&
      FillStyle       =   0  'Solid
      Height          =   4815
      Left            =   120
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' I did not comment that much as the reed solomon and galois field code is very very complicated
' Even I don't know how it works precisely
'

Private Type struct_files
  FileName As String
  Size     As Long
  index    As Long
End Type

Private busy As Boolean

Private Sub Encode()

  Dim i As Long
  Dim j As Long
  Dim n As Long
  Dim m As Long
  Dim rows As Long
  Dim cols As Long
  Dim fileo As String
  Dim tmp As Long
  Dim sz As Long
  Dim factors() As Long
  Dim vdm() As Long
  Dim factor As Long
  Dim blocksize As Long
  Dim fixhdr As struct_FIX_header
  Dim fileblock() As struct_File_Entry
  Dim cumufactors() As Long
  Dim h As Long
  Dim f As Boolean
  
  If initialize(8) = False Then
    MsgBox "Error initializing encoder/decoder", vbInformation Or vbOKOnly, App.Title
    Exit Sub
  End If
  If Len(Trim(fixprefix)) = 0 Then
    MsgBox "No fixfile prefix specified.", vbInformation Or vbOKOnly, App.Title
    Exit Sub
  End If
  n = list.ListItems.Count
  If n = 0 Then
    MsgBox "First select sourcefiles to encode by dragging and dropping them.", vbInformation Or vbOKOnly, App.Title
    Exit Sub
  End If
  If n < 2 Or n > 255 Then
    MsgBox "The number of sourcefiles must be in the range from 2 to 255.", vbInformation Or vbOKOnly, App.Title
    Exit Sub
  End If
  m = n * (Val(fixnumprc.Value) / 100)
  If m < 1 Then
    MsgBox "There must be at least 1 fixfile, please alter the 'Repair files to create' ratio.", vbInformation Or vbOKOnly, App.Title
    Exit Sub
  End If
  rows = n + m
  cols = n
   
  blocksize = 0
  For i = 1 To n
    If Val(list.ListItems(i).Tag) > blocksize Then
       blocksize = Val(list.ListItems(i).Tag)
    End If
  Next

  If blocksize > FIX_MAX_SIZE Then
    MsgBox "Sourcefiles may not be larger than " & FormatSize(FIX_MAX_SIZE), vbInformation Or vbOKOnly, App.Title
    Exit Sub
  End If
  
  ReDim factors(n - 1)
  ReDim fileblock(n - 1)
  ReDim dataParity(blocksize - 1)
  ReDim dataMul(blocksize - 1)
  ReDim cumufactors(m - 1, n - 1)
  
  ShowStatus ""
  ShowStatus "Started"
  ShowStatus "Collecting files ..."
  
  ShowInfo n, m, blocksize, 0, 0, 0
  
  For i = 0 To n - 1
    factors(i) = 1
    With fileblock(i)
      .FileName = list.ListItems(i + 1).Text
      .Size = FileLen(fixsavedir & "\" & Trim(.FileName))
      If MapFileMemory(fixsavedir & "\" & Trim(.FileName)) Then
        If chkfullCRC.Value = vbChecked Then
          .crc = getCRC32(.Size)
        Else
          h = .Size
          If h > FIX_SHORT_CRC_LENGTH Then
             h = FIX_SHORT_CRC_LENGTH
          End If
          .crc = getCRC32(h)
        End If
        UnMapFileMemory
      Else
        ShowStatus "Error opening " & Trim(.FileName) & " ..."
        Exit Sub
      End If
    End With
    DoEvents
  Next
  
  ShowStatus "Initializing encoder ..."
  
  vdm() = gf_make_dispersal_matrix(rows, cols)
  
  With fixhdr
   .sig = FIX_SIG
   .version = FIX_VER
   .datsz = blocksize
   .fixnum = m
   .orgnum = n
   .crc = &HFFFFFFFF
   .fullcrc = chkfullCRC.Value
  End With
  
  For i = 0 To rows - cols - 1
    For j = 0 To cols - 1
       cumufactors(i, j) = -1
    Next
  Next
  
  prg.Value = 0
  prg.Visible = True
  lblPrc.Visible = True
  Dim ss As Double
  
  For i = cols To rows - 1
     fileo = fixprefix & Format(i - cols, "_00#") & ".fix"
     ShowStatus "Building " & fileo & " ..."
     empBuffer VarPtr(dataParity(LBound(dataParity))), blocksize
     For j = 0 To cols - 1
        tmp = vdm(i, j)
        If tmp <> 0 Then
           factor = gf_single_divide(tmp, factors(j))
           factors(j) = tmp
           cumufactors(i - cols, j) = factor
           If MapFileMemory(fixsavedir & "\" & Trim(fileblock(j).FileName), fileblock(j).Size) Then
              empBuffer VarPtr(dataMul(LBound(dataMul))), blocksize
              f = True
              For h = 0 To i - cols
                If cumufactors(h, j) <> -1 Then
                  gf_mult_region IIf(f, fileblock(j).Size, blocksize), cumufactors(h, j), , Not f
                  f = False
                End If
              Next
              gf_add_parity blocksize
              UnMapFileMemory
           Else
              ShowStatus "Error opening " & Trim(fileblock(j).FileName) & " ..."
              prg.Visible = False
              lblPrc.Visible = False
              Exit Sub
           End If
        End If
        ss = ((i - cols) * cols + j) / ((rows - cols - 1) * cols + cols) * 100
        If ss < 0 Then ss = 0
        If ss > 100 Then ss = 100
        prg.Value = CInt(ss)
        lblPrc = Round(ss, 2) & "%"
        DoEvents
     Next
     Open fixsavedir & "\" & fileo For Binary As #1
        Put #1, , fixhdr
        Put #1, , fileblock()
        Put #1, , dataParity()
     Close #1
  Next
  ShowStatus "Finished"
  prg.Value = 100
  lblPrc = "100%"
   
  prg.Visible = False
  lblPrc.Visible = False
  
  DoEvents
  
End Sub

Private Sub Decode(Optional OnlyVerify As Boolean = False, Optional Cmd As String = vbNullString, Optional AfterCheck As Boolean = False)

  Dim rfile As Long
  Dim cfile As Long
  Dim mfile As Long
  Dim ffile As Long
  Dim gfile As Long
  Dim i As Long
  Dim j As Long
  Dim n As Long
  Dim m As Long
  Dim o As Long
  Dim h As Long
  Dim f As Boolean
  Dim rows As Long
  Dim cols As Long
  Dim fileo As String
  Dim path As String
  Dim tmp As Long
  Dim factors() As Long
  Dim vdm() As Long
  Dim factor As Long
  Dim blocksize As Long
  Dim exists() As Long
  Dim map() As Long
  Dim inv() As Long
  Dim rf As String
  Dim Offset As Long
  Dim xl As ListItem
  Dim valid As Boolean
  Dim fixhdr As struct_FIX_header
  Dim files() As struct_files
  Dim fileblock() As struct_File_Entry
  Dim cm As struc_Condensed_Matrix
  Dim presentCRCs As Collection
  
  If initialize(8) = False Then
    MsgBox "Error initializing encoder/decoder", vbInformation Or vbOKOnly, App.Title
    Exit Sub
  End If

  If Len(Cmd) = 0 Then
     For i = 1 To list.ListItems.Count
        With list.ListItems(i)
          GetBaseFilename .Text, fileo
          If StrComp(fileo, REAL_EXTENSION, vbTextCompare) = 0 Then
             fileo = fixsavedir & "\" & .Text
             Me.Tag = fileo
             Exit For
          End If
        End With
        fileo = ""
     Next
     If Len(Me.Tag) = 0 Then
        MsgBox "No fixfiles found in the current list, verify/decoding aborted.", vbInformation Or vbOKOnly, App.Title
        Exit Sub
     End If
     fileo = Me.Tag
  Else
     fileo = Cmd
  End If
  
  On Error GoTo BadFilename
  If Len(Dir(fileo)) = 0 Then
     MsgBox "File could not be found. (" & fileo & ")", vbOKOnly Or vbInformation, App.Title
     Close #1
     Exit Sub
  End If
  On Error GoTo 0
 
  list.ListItems.Clear
  
  ShowStatus ""
  ShowStatus "Started"
  ShowStatus "Checking file ..."
    
  Open fileo For Binary As #1
     Get #1, , fixhdr
      
     If fixhdr.sig <> FIX_SIG Then
        MsgBox "No FIX file selected. (" & fileo & ")", vbOKOnly Or vbInformation
        Close #1
        Exit Sub
     End If
     If fixhdr.version > FIX_VER Then
        MsgBox "The FIX file has a unsupported version.", vbOKOnly Or vbInformation
        Close #1
        Exit Sub
     End If
     
     n = fixhdr.orgnum
     m = fixhdr.fixnum
     rows = n + m
     cols = n
     blocksize = fixhdr.datsz
  
     If n > FIX_MAX_DATA Or m > FIX_MAX_DATA Then
        MsgBox "FIX header corrupt.", vbOKOnly Or vbInformation
        Close #1
        Exit Sub
     End If
  
     ReDim exists(rows - 1)
     ReDim factors(rows - 1)
     ReDim map(rows - 1)
     ReDim fileblock(n - 1)
     ReDim files(rows - 1)
     ReDim dataMul(blocksize - 1)
     ReDim dataParity(blocksize - 1)
     
     Get #1, , fileblock()
  Close #1
  
  GetFilename fileo, path
  fixsavedir = path
  chkfullCRC.Value = IIf(fixhdr.fullcrc, vbChecked, vbUnchecked)
        
  ShowStatus "Collecting files ..."
          
  h = 0
  Set presentCRCs = New Collection
  Dim pCRC As IUnknown
  
  'map CRCs for each file in the working directory
  fileo = Dir(path & "\*.*")
  Do While Len(fileo) > 0
    If MapFileMemory(path & "\" & fileo) Then
       h = UBound(dataBuffer) + 1
       If Not fixhdr.fullcrc Then
          If h > FIX_SHORT_CRC_LENGTH Then
             h = FIX_SHORT_CRC_LENGTH
          End If
       End If
       presentCRCs.add fileo & ":" & getCRC32(h)
    UnMapFileMemory
    End If
    fileo = Dir()
  Loop
  
  'On Error GoTo BadFilename
  
  j = 0
  h = 0
  gfile = 0   'good files
  rfile = 0   'renamed files
  cfile = 0   'corrupt files
  mfile = 0   'missing files
  ffile = 0   'fix files
  For i = 0 To rows - 1
    If j >= cols Then Exit For
    If i < cols Then
       fileo = fixsavedir & "\" & Trim(fileblock(i).FileName)
    Else
       fileo = fixsavedir & "\" & Format(i - cols, "\*_00#\.fix")
    End If
    rf = Dir(fileo)
    If Len(rf) = 0 Then
    
       valid = False
       'search mapped CRCs if the file is renamed
       'we do this by checking the crc of each partial file
       'and compare it with the stored crc
       If i < cols Then
         For o = 1 To presentCRCs.Count
           If Val(Split(presentCRCs(o), ":")(1)) = fileblock(i).crc Then
              If Not OnlyVerify Then
                 Name fixsavedir & "\" & Split(presentCRCs(o), ":")(0) As fileo
              Else
                 fileo = fixsavedir & "\" & Split(presentCRCs(o), ":")(0)
              End If
              With list.ListItems
                Set xl = .add(, , GetFilename(fileo, path))
                xl.Tag = FileLen(fileo)
                xl.SubItems(1) = FormatSize(xl.Tag)
                xl.SubItems(2) = "Renamed"
                xl.ListSubItems(2).ForeColor = vbRed
                xl.ListSubItems(2).Bold = True
                xl.ForeColor = vbRed
                xl.EnsureVisible
                rfile = rfile + 1
              End With
              With files(j)
                .FileName = fileo
                .Size = Val(xl.Tag)
                .index = i
              End With
              map(i) = j
              j = j + 1
              valid = True
              Exit For
           End If
         Next
       End If
       
       'file does not exists
       If Not valid Then
         map(i) = -1
         'exclude missing .fix file from the list by just skipping them
         If i < cols Then
           With list.ListItems
             Set xl = .add(, , GetFilename(fileo, path))
             xl.Tag = Format(0)
             xl.SubItems(1) = ""
             xl.SubItems(2) = "Missing"
             xl.ListSubItems(2).ForeColor = vbRed
             xl.ListSubItems(2).Bold = True
             xl.ForeColor = vbRed
             xl.EnsureVisible
             mfile = mfile + 1
           End With
         End If
      End If
    Else
       'file exists
       'if it is a original file, check the crc to see if it is valid
       'if not, we mark it as not present so it will be rebuild
       valid = True
       
       If i < cols Then
         If MapFileMemory(fixsavedir & "\" & rf) Then
            h = fileblock(i).Size
            If Not fixhdr.fullcrc Then
              If h > FIX_SHORT_CRC_LENGTH Then
                 h = FIX_SHORT_CRC_LENGTH
              End If
            End If
            If getCRC32(h) <> fileblock(i).crc Then
               valid = False
            Else
               gfile = gfile + 1
            End If
            UnMapFileMemory
         Else
            ShowStatus "Error opening " & Trim(fileblock(j).FileName) & " ..."
            Exit Sub
         End If
       End If
       
       If valid Then
         With files(j)
           .FileName = fixsavedir & "\" & rf
           .Size = FileLen(.FileName)
           .index = i
         End With
       
         With list.ListItems
           Set xl = .add(, , GetFilename(rf, path))
           xl.Tag = Format(files(j).Size)
           xl.SubItems(1) = FormatSize(files(j).Size)
           xl.SubItems(2) = "Good"
           xl.ListSubItems(2).ForeColor = vbGreen
           xl.ListSubItems(2).Bold = True
           If i >= cols Then
             ffile = ffile + 1
             xl.ForeColor = vbBlue
           End If
         End With
       
         map(i) = j
         j = j + 1
       Else
         map(i) = -1
         
         With list.ListItems
           Set xl = .add(, , GetFilename(rf, path))
           xl.Tag = Format(0)
           xl.SubItems(1) = ""
           xl.SubItems(2) = "Corrupt"
           xl.ListSubItems(2).ForeColor = vbRed
           xl.ListSubItems(2).Bold = True
           xl.EnsureVisible
           cfile = cfile + 1
         End With
       
       End If
    End If
    DoEvents
  Next
  
  Set presentCRCs = Nothing
  
  ShowInfo n, ffile, blocksize, cfile, mfile, rfile
  
  If OnlyVerify Then
    ShowStatus "Finished"
    If AfterCheck Then
       If gfile = n Then
         MsgBox "Repair succeeded.", vbInformation Or vbOKOnly, App.Title
       Else
         MsgBox "Some files could not be repaired.", vbInformation Or vbOKOnly, App.Title
       End If
    Else
       If gfile = n Then
         MsgBox "Repair not needed.", vbInformation Or vbOKOnly, App.Title
       Else
         MsgBox "Some files need to be repaired.", vbInformation Or vbOKOnly, App.Title
       End If
    End If
    Exit Sub
  End If

  If gfile = n Then
     ShowStatus "Finished"
     MsgBox "Repair not needed.", vbInformation Or vbOKOnly, App.Title
     Exit Sub
  End If
  
  'aantal datafiles mag niet kleiner zijn dan het aantal headerfiles - aantal fixfiles
  If j < n Then
     ShowStatus "Not enough fixfiles to repair ..."
     MsgBox "Need " & n - j + ffile & " blocks to repair, only " & ffile & " found.", vbInformation Or vbOKOnly, App.Title
     Exit Sub
  End If
  
  ShowStatus "Initializing decoder ..."
  
  For i = 0 To rows - 1
    exists(i) = IIf(map(i) <> -1, True, False)
    factors(i) = 1
  Next
 
  vdm() = gf_make_dispersal_matrix(rows, cols)
  cm = gf_condense_dispersal_matrix(vdm(), exists(), rows, cols)
    
  For i = 0 To cols - 1
    If map(i) = -1 Then map(i) = map(cm.row_identities(i))
  Next
  
  inv() = gf_invert_matrix(cm.condensed_mat(), cols)
   
  Dim cumufactors() As Long
  ReDim cumufactors(cols - 1, cols - 1)
  
  For i = 0 To cols - 1
    For j = 0 To cols - 1
       cumufactors(i, j) = -1
    Next
  Next
   
  i = 0
  prg.Value = 0
  prg.Visible = True
  lblPrc.Visible = True
  Dim ss As Double
  ss = 0
  ffile = 1
  Do While i < cols
     If cm.row_identities(i) < cols Then
        If factors(i) <> 1 Then
           tmp = gf_single_divide(1, factors(i))
           factors(i) = 1
           cumufactors(i, map(i)) = tmp
        End If
     Else
        fileo = fixsavedir & "\" & Trim(fileblock(i).FileName)
        ShowStatus "Repairing " & GetFilename(fileo, vbNullString) & " ..."
        empBuffer VarPtr(dataParity(LBound(dataParity))), blocksize
        For j = 0 To cols - 1
           tmp = inv(i, j)
           factor = gf_single_divide(tmp, factors(j))
           factors(j) = tmp
           cumufactors(i, map(j)) = factor
           If files(map(j)).index >= n Then
              Offset = Len(fixhdr) + Len(fileblock(0)) * fixhdr.orgnum
           Else
              Offset = 0
           End If
           If MapFileMemory(files(map(j)).FileName, files(map(j)).Size + Offset) Then
              empBuffer VarPtr(dataMul(LBound(dataMul))), blocksize
              f = True
              For h = 0 To i
                If cumufactors(h, map(j)) <> -1 Then
                  gf_mult_region IIf(f, files(map(j)).Size, blocksize), cumufactors(h, map(j)), Offset, Not f
                  f = False
                End If
              Next
              gf_add_parity blocksize
              UnMapFileMemory
           Else
              ShowStatus "Error opening " & Trim(fileblock(j).FileName) & " ..."
              prg.Visible = False
              lblPrc.Visible = False
              Exit Sub
           End If
           ss = (ffile * cols + j) / ((mfile + cfile) * cols + cols) * 100
           If ss < 0 Then ss = 0
           If ss > 100 Then ss = 100
           prg.Value = CInt(ss)
           lblPrc = Round(ss, 2) & "%"
           DoEvents
       Next
       ffile = ffile + 1
       Open fileo For Binary As #1
       Put #1, , dataParity()
       Close #1
       'truncate the file to the original length
       If Not TruncateFile(fileo, fileblock(i).Size) Then
          ShowStatus "Could not truncate " & Trim(fileblock(j).FileName) & " to it's original length..."
          prg.Visible = False
          lblPrc.Visible = False
          Exit Sub
       End If
     End If
     i = i + 1
  Loop
  ShowStatus "Finished"
  prg.Value = 100
  lblPrc = "100%"
  
  prg.Visible = False
  lblPrc.Visible = False
  
  On Error GoTo 0
  
  DoEvents
  
  're-itterate to see if repair was succesful
  Decode True, Cmd, True
  
Exit Sub
BadFilename:
  MsgBox "The specified file could not be openend.", vbInformation Or vbOKOnly, App.Title
  On Error GoTo 0
End Sub

Private Sub fixnumprc_Change()
  lblfixnumprcval = Format(fixnumprc.Value, "#\%")
End Sub

Private Sub Form_Load()
  Dim Cmd As String
  Dim Vo As Boolean
  
  Me.Caption = REAL_APP_NAME
  StartED.Caption = "Encode"
  
  prg.Min = 0
  prg.Max = 100
  fixnumprc.Min = 5
  fixnumprc.Max = 100
  fixnumprc.SmallChange = 5
  fixnumprc.LargeChange = 10
  ' init on 20% reapair files
  fixnumprc.Value = 20
  
  prg.Visible = False
  lblPrc.Visible = False
  Verify.Enabled = False
  
  ShowStatus vbNullString
  ShowInfo
  
  NewED_Click
  
  If GetSetting(App.Title, REAL_APP_NAME, "AssociatedFiles", False) = False Then
    If MsgBox("Would you like to associate this program with " & REAL_EXTENSION & " repair files?", vbYesNo Or vbQuestion, App.Title) = vbYes Then
       Associate App.path & "\" & App.EXEName, REAL_EXTENSION, REAL_CONTEXT_NAME, App.path & "\" & App.EXEName & ".exe,2", "%1"
       CreateNewKey HKEY_CLASSES_ROOT, "*\shell\" & REAL_CONTEXT_ACTION_NAME_VERIFY & "\Command"
       SetKeyValue HKEY_CLASSES_ROOT, "*\shell\" & REAL_CONTEXT_ACTION_NAME_VERIFY & "\Command", "", App.path & "\" & App.EXEName & ".exe " & CMD_VERIFY_ONLY & " %1", REG_SZ
       CreateNewKey HKEY_CLASSES_ROOT, "*\shell\" & REAL_CONTEXT_ACTION_NAME_REPAIR & "\Command"
       SetKeyValue HKEY_CLASSES_ROOT, "*\shell\" & REAL_CONTEXT_ACTION_NAME_REPAIR & "\Command", "", App.path & "\" & App.EXEName & ".exe %1", REG_SZ
       SaveSetting App.Title, REAL_APP_NAME, "AssociatedFiles", True
    End If
  End If
  
  If Len(Command) Then
    busy = True
    mode(1).Value = True
    busy = False
    
    Cmd = Command
    
    If InStr(1, Cmd, CMD_REMOVE_ASSOC, vbTextCompare) > 0 Then
       RemoveAssociate App.path & "\" & App.EXEName, REAL_EXTENSION
       SaveSetting App.Title, REAL_APP_NAME, "AssociatedFiles", False
       Form_Unload 0
    End If
    
    Vo = False
    If InStr(1, Cmd, CMD_VERIFY_ONLY, vbTextCompare) > 0 Then
       Vo = True
       Cmd = Replace(Cmd, CMD_VERIFY_ONLY, vbNullString, , , vbTextCompare)
    End If
    
    Me.Show
    DoEvents
    Decode Vo, Trim(Cmd)
  Else
    busy = True
    mode(0).Value = True
    busy = False
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
  End
End Sub

Private Sub list_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then
     If Not list.SelectedItem Is Nothing Then
       list.ListItems.Remove (list.SelectedItem.index)
       If Not list.SelectedItem Is Nothing Then list.SelectedItem.Selected = True
     End If
   End If
End Sub

Private Sub list_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  If mode(0).Value = True Then
    Dim path As String
    Dim ext As String
    Dim i As Long
    Dim S As Long
    Dim xl As ListItem
    With list.ListItems
      For i = 1 To data.files.Count
        Set xl = .add(, , GetFilename(data.files.Item(i), path))
        If Len(fixsavedir) = 0 Then
           fixsavedir = path
        End If
        S = FileLen(path & "\" & xl.Text)
        xl.Tag = Format(S)
        xl.SubItems(1) = FormatSize(S)
        xl.SubItems(2) = "Unverified"
      Next
      If .Count > 0 Then
        fixprefix = GetBaseFilename(.Item(1).Text, vbNullString)
      End If
    End With
  Else
    With list.ListItems
      If data.files.Count > 1 Or list.ListItems.Count = 1 Then
         MsgBox "In decode mode you only have to drag and drop 1 of the " & REAL_EXTENSION & " repair files.", vbInformation Or vbOKOnly, App.Title
         Exit Sub
      End If
      For i = 1 To data.files.Count
        GetBaseFilename data.files.Item(i), ext
        If StrComp(ext, REAL_EXTENSION, vbTextCompare) = 0 Then
          Set xl = .add(, , GetFilename(data.files.Item(i), path))
          If Len(fixsavedir) = 0 Then
            fixsavedir = path
          End If
          S = FileLen(path & "\" & xl.Text)
          xl.Tag = Format(S)
          xl.SubItems(1) = FormatSize(S)
          xl.SubItems(2) = "Unverified"
        Else
          MsgBox "In decode mode you only have to drag and drop 1 of the " & REAL_EXTENSION & " repair files.", vbInformation Or vbOKOnly, App.Title
          Exit Sub
        End If
      Next
      If .Count > 0 Then
        fixprefix = GetBaseFilename(.Item(1).Text, vbNullString)
      End If
    End With
  End If
End Sub

Private Sub mode_Click(index As Integer)
  If Not busy Then
    If MsgBox("This action clears the current settings, do you want to continue?", vbInformation Or vbYesNo, App.Title) = vbNo Then
       busy = True
       mode(IIf(index = 0, 1, 0)).Value = True
       busy = False
       Exit Sub
    End If
  End If
  NewED_Click
  If mode(0).Value = True Then
     StartED.Caption = "Encode"
     lblfixnumprc.Enabled = True
     lblfixsavedir.Enabled = False
     lblfixprefix.Enabled = True
     lblfixnumprcval.Enabled = True
     fixnumprc.Enabled = True
     fixsavedir.Enabled = False
     fixprefix.Enabled = True
     Verify.Enabled = False
  Else
     StartED.Caption = "Decode"
     lblfixnumprc.Enabled = False
     lblfixsavedir.Enabled = False
     lblfixprefix.Enabled = False
     lblfixnumprcval.Enabled = False
     fixnumprc.Enabled = False
     fixsavedir.Enabled = False
     fixprefix.Enabled = False
     Verify.Enabled = True
  End If
End Sub

Private Sub NewED_Click()
  list.ListItems.Clear
  ShowStatus vbNullString
  ShowInfo
  fixsavedir = vbNullString
  fixprefix = "New Fix Set"
End Sub

Private Sub StartED_Click()
  StartED.Enabled = False
  NewED.Enabled = False
  If StartED.Caption = "Encode" Then
    Encode
  Else
    Decode
  End If
  StartED.Enabled = True
  NewED.Enabled = True
End Sub

Private Sub ShowStatus(Message As String)
  If Len(Message) = 0 Then
    lstStatus.Clear
  Else
    lstStatus.AddItem Date & " " & Time() & "> " & Message
    lstStatus.ListIndex = lstStatus.ListCount - 1
  End If
End Sub

Private Sub ShowInfo(Optional sf As Long = -1, Optional ff As Long = -1, Optional bs As Long = -1, Optional df As Long = -1, Optional mf As Long = -1, Optional rf As Long = -1)
  Dim i As Long
  i = 0
  If sf <> -1 Then i = i + 1: lblInfo(0) = sf
  If ff <> -1 Then i = i + 1: lblInfo(1) = ff
  If bs <> -1 Then i = i + 1: lblInfo(2) = FormatSize(bs)
  If mf <> -1 Then i = i + 1: lblInfo(3) = mf
  If rf <> -1 Then i = i + 1: lblInfo(4) = rf
  If df <> -1 Then i = i + 1: lblInfo(5) = df
  If i = 0 Then
    For i = 0 To 5
       lblInfo(i) = 0
       lblInfo(i).ForeColor = vbWhite
    Next
  Else
    For i = 3 To 5
       lblInfo(i).ForeColor = IIf(Val(lblInfo(i)) = 0, vbGreen, vbRed)
    Next
  End If
End Sub

Private Sub Verify_Click()
  Decode True, vbNullString
End Sub
