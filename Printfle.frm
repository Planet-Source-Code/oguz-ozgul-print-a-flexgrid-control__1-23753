VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPrintFlex 
   Caption         =   "Print table"
   ClientHeight    =   5280
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkBorders 
      BackColor       =   &H8000000C&
      Caption         =   "Borders"
      Height          =   195
      Left            =   2190
      TabIndex        =   33
      Top             =   870
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Page orientation"
      Height          =   1455
      Left            =   0
      TabIndex        =   30
      Top             =   3120
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Portrait"
         Height          =   255
         Left            =   690
         TabIndex        =   32
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Landscape"
         Height          =   255
         Left            =   690
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   165
         Picture         =   "Printfle.frx":0000
         Top             =   360
         Width           =   390
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   120
         Picture         =   "Printfle.frx":0A42
         Top             =   885
         Width           =   480
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
      Height          =   600
      Left            =   2190
      TabIndex        =   29
      ToolTipText     =   "Do not print and exits"
      Top             =   4680
      Width           =   1605
   End
   Begin VB.ComboBox CmbFontSize 
      Height          =   315
      ItemData        =   "Printfle.frx":1444
      Left            =   5760
      List            =   "Printfle.frx":1475
      TabIndex        =   8
      Text            =   "CmbFontSize"
      ToolTipText     =   "Font size will be set to the nearest available value"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton CmdMore 
      Caption         =   "More..."
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Set more properties like page size etc."
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox CmbPrinters 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Combo1"
      ToolTipText     =   "System printers list"
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Picture         =   "Printfle.frx":14B0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Re-count the table size with the new font and margin settings"
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox ChkPageNumbers 
      BackColor       =   &H8000000C&
      Caption         =   "Page numbers"
      Height          =   195
      Left            =   4170
      TabIndex        =   6
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.VScrollBar ValPageId 
      Height          =   495
      Left            =   6360
      Min             =   1
      TabIndex        =   7
      Top             =   960
      Value           =   1
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Margins"
      Height          =   2295
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   2055
      Begin VB.TextBox TxtBottomMargin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   780
         TabIndex        =   5
         Text            =   "20"
         Top             =   1620
         Width           =   495
      End
      Begin VB.TextBox TxtTopMargin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   780
         TabIndex        =   2
         Text            =   "20"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TxtRightMargin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "20"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox TxtLeftMargin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "20"
         Top             =   960
         Width           =   495
      End
      Begin VB.Image TmpImage 
         Height          =   480
         Left            =   840
         Picture         =   "Printfle.frx":3232
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label9 
         Caption         =   "Top"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   900
         TabIndex        =   28
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1530
         TabIndex        =   27
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   240
         TabIndex        =   26
         Top             =   810
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Bottom"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   810
         TabIndex        =   25
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Between 10 - 30 mm"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1980
         Width           =   1455
      End
   End
   Begin VB.PictureBox PicPreview 
      BackColor       =   &H8000000C&
      Height          =   3375
      Left            =   2160
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   12
      Top             =   840
      Width           =   3375
      Begin VB.Shape ShpTable 
         BackColor       =   &H8000000A&
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   720
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LblTableText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Table created by WinMrp2000, Page : 1 / 4"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   3
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label LblPageNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   2400
         TabIndex        =   17
         Top             =   2760
         Width           =   195
      End
      Begin VB.Shape ShpMargins 
         BorderColor     =   &H8000000B&
         Height          =   2535
         Left            =   720
         Top             =   360
         Width           =   1935
      End
      Begin VB.Shape ShpPage 
         BackStyle       =   1  'Opaque
         Height          =   2775
         Left            =   600
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3780
      Picture         =   "Printfle.frx":3C74
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Prints the table to the selected printer"
      Top             =   4680
      Width           =   2910
   End
   Begin MSFlexGridLib.MSFlexGrid Fg2 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   4920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      _Version        =   393216
      FormatString    =   ""
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   3840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label LblTotalPages 
      Caption         =   "Label6"
      Height          =   255
      Left            =   5640
      TabIndex        =   24
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Font size"
      Height          =   255
      Left            =   5760
      TabIndex        =   23
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Select printer"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label LblHeightWidth 
      Caption         =   "290 x 210 mm"
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Page"
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.Label LblPageId 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      ToolTipText     =   "Current page"
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Preview"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "FrmPrintFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OxGP As Ox_TableProp
Dim PageHWRate As Single
Dim LeftMarginPageRate As Single
Dim RightMarginPageRate As Single
Dim TopMarginPageRate As Single
Dim BottomMarginPageRate As Single
Dim Fg1 As MSFlexGrid
Dim FgContainer As Form
Dim FormatStr As String
Dim m As Ox_Margins
Dim CurrentPage As Long
Dim DefPrinter As Printer
Dim HeaderRows As Long, FooterRows As Long

Private Sub ChkPageNumbers_Click()
    LblPageNo.Visible = Not LblPageNo.Visible
End Sub

Private Sub CmbPrinters_Change()
    Dim APrinter As Printer
    For Each APrinter In Printers
        If APrinter.DeviceName = CmbPrinters.Text Then
            Set DefPrinter = Printers(CmbPrinters.ListIndex)
        End If
    Next APrinter
End Sub

Private Sub CmdMore_Click()
    Cd1.ShowPrinter
End Sub

Private Sub Command1_Click()
    Dim MsgBoxResult As Long
    
    If ShpTable.Width > ShpMargins.Width Then
        MsgBoxResult = MsgBox("Table width is out of the margins. Click Ok to continue anyway, or click cancel and set margins, or try to increase the font size, or resize the table by hand to make it fit.", vbOKCancel, "!")
        If MsgBoxResult = 2 Then Exit Sub
    End If
    
    If ShpTable.Top + ShpTable.Height > ShpMargins.Top + ShpMargins.Height Then
        MsgBoxResult = MsgBox("Click OK to print, or click Cancel, then click REFRESH button to see a preview of the exact result of printing.", vbOKCancel, "Page settings changed, print preview update")
        If MsgBoxResult = 2 Then Exit Sub
    End If
    
    OxGP = GeneratePreview(FgContainer, Fg1, Fg2, Cd1, HeaderRows, FooterRows, FormatStr, m, Val(CmbFontSize.Text))
    QuickPrintFlexGrid FgContainer, Fg1, Fg2, Cd1, 2, 3, "", m, ChkPageNumbers.Value, ChkBorders.Value
End Sub

Public Function TransferValues(ByVal TrnFgContainer As Form, ByVal TrnFg As MSFlexGrid, ByVal TrnHeaderRows As Long, ByVal TrnFooterRows As Long, ByVal TrnFormatString As String)
    Set FgContainer = TrnFgContainer
    Set Fg1 = TrnFg
    HeaderRows = TrnHeaderRows
    FooterRows = TrnFooterRows
    FormatStr = TrnFormatString
End Function

Private Sub Command2_Click()
    m.BottomInmm = Val(TxtBottomMargin.Text)
    m.LeftInmm = Val(TxtLeftMargin.Text)
    m.RightInmm = Val(TxtRightMargin.Text)
    m.TopInmm = Val(TxtTopMargin.Text)
    OxGP = GeneratePreview(FgContainer, Fg1, Fg2, Cd1, HeaderRows, FooterRows, FormatStr, m, Val(CmbFontSize.Text))
    ValPageId.Value = 1
    LblHeightWidth.Caption = CStr(Int(TwipsTomm(Printer.Height))) & " x " & CStr(Int(TwipsTomm(Printer.Width))) & " cm"
    SetShapes 1
End Sub

Private Sub Form_Load()
    
    Dim tmp As Long
    Dim APrinter As Printer
    
    Select Case Printer.Orientation
        Case vbPRORLandscape
            Option2.Value = True
        Case Else
            Option1.Value = True
    End Select
    Me.ScaleMode = vbPixels
    Picture1.ScaleMode = vbPixels
    tmp = CreateEllipticRgn(0, 0, PicPreview.Width, PicPreview.Height)
    SetWindowRgn Picture1.hWnd, tmp, False
    Me.ScaleMode = vbTwips
    Picture1.ScaleMode = vbTwips
    
    Set FlexGrid = Fg1
    For Each APrinter In Printers
        CmbPrinters.AddItem APrinter.DeviceName
        If APrinter.DeviceName = Printer.DeviceName Then
            CmbPrinters.ListIndex = CmbPrinters.ListCount - 1
        End If
    Next APrinter
    m.BottomInmm = Val(TxtBottomMargin.Text)
    m.LeftInmm = Val(TxtLeftMargin.Text)
    m.RightInmm = Val(TxtRightMargin.Text)
    m.TopInmm = Val(TxtTopMargin.Text)
    CmbFontSize.Text = Closest(Printer.FontSize)
    OxGP = GeneratePreview(FgContainer, Fg1, Fg2, Cd1, HeaderRows, FooterRows, FormatStr, m, Val(CmbFontSize.Text))
    ValPageId.Value = 1
    SetShapes 1
End Sub

Private Sub SetShapes(ByVal PageId As Long)

    PageHWRate = Printer.Height / Printer.Width
    TableHWRate = OxGP.TableHeights(1) / OxGP.TableWidth
    LeftMarginPageRate = Val(TxtLeftMargin.Text) / TwipsTomm(Printer.Width)
    RightMarginPageRate = Val(TxtRightMargin.Text) / TwipsTomm(Printer.Width)
    TopMarginPageRate = Val(TxtTopMargin.Text) / TwipsTomm(Printer.Height)
    BottomMarginPageRate = Val(TxtBottomMargin.Text) / TwipsTomm(Printer.Height)
    If PageHWRate <= 1 Then 'Set the width first, then fit the width
        ShpPage.Width = 5 * (PicPreview.Width / 6) '5/6 of the picture box
        ShpPage.Height = ShpPage.Width * PageHWRate
    Else 'Set the height first, then fit the width
        ShpPage.Height = 5 * (PicPreview.Height / 6) '5/6 of the picture box
        ShpPage.Width = ShpPage.Height / PageHWRate
    End If
    ShpPage.Left = (PicPreview.Width - ShpPage.Width) / 2
    ShpPage.Top = (PicPreview.Height - ShpPage.Height) / 2
    ShpMargins.Top = ShpPage.Top + ShpPage.Height * (TopMarginPageRate)
    ShpMargins.Left = ShpPage.Left + ShpPage.Width * LeftMarginPageRate
    ShpMargins.Width = ShpPage.Width * (1 - LeftMarginPageRate - RightMarginPageRate)
    ShpMargins.Height = ShpPage.Height * (1 - TopMarginPageRate - BottomMarginPageRate)
    ShpTable.Left = ShpMargins.Left
    ShpTable.Top = ShpMargins.Top + ShpPage.Height * (mmToTwips(2 * Printer.TextHeight("table cre")) / Printer.Height)
    ShpTable.Width = ShpPage.Width * (OxGP.TableWidth / TwipsTomm(Printer.Width))
    ShpTable.Height = ShpPage.Height * (OxGP.TableHeights(PageId) / TwipsTomm(Printer.Height))
    LblPageNo.Left = ShpMargins.Left + ShpMargins.Width - LblPageNo.Width
    LblPageNo.Top = ShpMargins.Top + ShpMargins.Height - LblTableText.Height / 2
    LblTableText.Top = ShpMargins.Top
    LblTableText.Left = ShpMargins.Left
    LblHeightWidth.Caption = CStr(Closest(TwipsTomm(Printer.Height))) & " x " & CStr(Closest(TwipsTomm(Printer.Width))) & " mm"
    LblTotalPages.Caption = "(" & CStr(OxGP.NumOfPages) & ") pages"
End Sub


Private Sub Option1_Click()
    TmpImage.Picture = Image1.Picture
    Printer.Orientation = vbPRORPortrait
    SetShapes CurrentPage
End Sub

Private Sub Option2_Click()
    TmpImage.Picture = Image2.Picture
    Printer.Orientation = vbPRORLandscape
    SetShapes CurrentPage
End Sub


Private Sub TxtBottomMargin_LostFocus()
    With TxtBottomMargin
        If Val(.Text) > 30 Then .Text = "30"
        If Val(.Text) < 10 Then .Text = "10"
        .Text = CStr(Val(.Text))
    End With
    SetShapes CurrentPage
End Sub

Private Sub TxtHeaderRows_lostfocus()
    Dim tmp As Long, tmp2 As Long
    
    With TxtHeaderRows
        tmp2 = Val(.Text)
        tmp = tmp2
        If tmp2 <= 0 Then
            tmp = 0
        ElseIf tmp2 > Fg2.Rows - Val(TxtFooterRows.Text) Then
            tmp = Fg2.Rows - Val(TxtFooterRows.Text)
        End If
    End With
    TxtHeaderRows.Text = CStr(tmp)
End Sub

Private Sub TxtLeftMargin_LostFocus()
    With TxtLeftMargin
        If Val(.Text) > 30 Then .Text = "30"
        If Val(.Text) < 10 Then .Text = "10"
        .Text = CStr(Val(.Text))
    End With
    SetShapes CurrentPage
End Sub

Private Sub TxtRightMargin_LostFocus()
    With TxtRightMargin
        If Val(.Text) > 30 Then .Text = "30"
        If Val(.Text) < 10 Then .Text = "10"
        .Text = CStr(Val(.Text))
    End With
    SetShapes CurrentPage
End Sub

Private Sub TxtTopMargin_LostFocus()
    With TxtTopMargin
        If Val(.Text) > 30 Then .Text = "30"
        If Val(.Text) < 10 Then .Text = "10"
        .Text = CStr(Val(.Text))
    End With
    SetShapes CurrentPage
End Sub

Private Sub ValPageId_Change()
    If ValPageId.Value = OxGP.NumOfPages + 1 Then
        ValPageId.Value = OxGP.NumOfPages
        LblPageNo.Caption = CStr(ValPageId.Value)
        Exit Sub
    End If
    LblPageId.Caption = CStr(ValPageId.Value)
    LblPageNo.Caption = CStr(ValPageId.Value)
    CurrentPage = ValPageId.Value
    SetShapes CurrentPage
End Sub

Function Closest(ByVal Value As Single) As Long

    Dim upV As Long, downV As Long
    
    upV = Int(Value + 1)
    downV = Int(Value)
    
    If Abs(Value - upV) > Abs(Value - downV) Then 'It is closer to int value..
        Closest = downV
    Else
        Closest = upV
    End If
End Function
