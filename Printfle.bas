Attribute VB_Name = "PrintFlex"
Option Explicit

Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


'Type declerations

Type Ox_TableProp
    TableWidth As Single
    NumOfPages As Long
    TableHeights() As Single
End Type

Type Ox_Cell
    RowId As Long
    ColId As Long
End Type

Type Ox_Margins
    LeftInmm As Long
    RightInmm As Long
    TopInmm As Long
    BottomInmm As Long
End Type

Type Ox_ParseResult
    GroupCount As Long
    Groups() As String
    NewWidthInmm As Long
    NewHeightInmm As Long
    x As Single
    y As Single
    Defaultx As Single
    Defaulty As Single
    Alignment As Single
    BackColor As Long
End Type

Type Ox_UserFormat
    nOfColsToBeFormatted As Long
    nOfRowsToBeFormatted As Long
    ColsToBeFormatted() As Long
    RowsToBeFormatted() As Long
    ColWidth() As Long
    RowHeight() As Long
End Type

Type Ox_PageInfo
    StartRow As Long
    nRows As Long
    PageHeightInmm As Long
End Type

'Variable declerations

Dim i, j, k As Long
Dim NumOfPages As Long
Dim PageInfos() As Ox_PageInfo
Dim PrWInmm As Long, PrHInmm As Long, PrFontSize As Single
Dim PrHWithoutHAndF As Long 'Page height without header and footer
Dim HeaderHeightInmm As Long, FooterHeightInmm As Long
Dim FGWInmm As Single, FGHInmm As Single
Dim nHeaderRows As Long, nFooterRows As Long
Dim CellInfos() As Ox_ParseResult
Dim MaxWidthOfColInmm() As Long, MaxHeightOfRowInmm() As Long
Dim UserFormats() As Ox_UserFormat
Dim FlexGrid As MSFlexGrid
Dim FlexGrid2 As MSFlexGrid
Dim Margins As Ox_Margins
Dim CurrentPageFooterStartPosition As Single
Dim TableTextHeight As Single
Dim PrintToPicBox As Boolean
Dim PrintPageNo As Boolean
Dim TableBorders As Boolean
Dim MaxWidthWordInColumns() As Single

'Temporary variable declerations

Dim TmpCellWInmm As Long, TmpCellHInmm As Long
Dim TmpTextWInmm As Long, TmpTextHInmm As Long
Dim TmpScaleMode As Long
Dim TmpRow As Long, TmpCol As Long


'Functions and subroutines

Function CutLastWord(ByVal TextToCut As String)
'Returns the string back, without the last word
    Dim CLW_WordCount As Long, CLW_i As Long

    CLW_WordCount = WordsIn(TextToCut)
    For CLW_i = 1 To CLW_WordCount - 1
        CutLastWord = CutLastWord & nthWordOf(TextToCut, CLW_i) & " "
    Next CLW_i
    CutLastWord = RTrim(CutLastWord) 'Cuts the space at from the end

End Function

Function CountAlignment(ByVal CCA_Value As Long) As String
    Select Case CCA_Value
        Case 0
            CountAlignment = "Left"
        Case 1
            CountAlignment = "Left"
        Case 2
            CountAlignment = "Left"
        Case 3
            CountAlignment = "Center"
        Case 4
            CountAlignment = "Center"
        Case 5
            CountAlignment = "Center"
        Case 6
            CountAlignment = "Right"
        Case 7
            CountAlignment = "Right"
        Case 8
            CountAlignment = "Right"
        Case 9
            'This is left for strings and right for numbers..
            'But most of the time the value will be a number
            CountAlignment = "Right"
    End Select
End Function

Function WordsIn(ByVal TextToCount As String) As Long
'Returns the word count of a text
    Dim WI_i As Long, WI_Tmp As Long
    
    WordsIn = 1
    For WI_i = 1 To Len(TextToCount)
        WI_Tmp = InStr(TextToCount, " ")
        If WI_Tmp > 0 Then
            WordsIn = WordsIn + 1
            Mid(TextToCount, WI_Tmp, 1) = "-" 'For avoiding duplication
        Else
            Exit Function
        End If
    Next WI_i
End Function

Function nthWordOf(ByVal TextToSeperate As String, ByVal n As Long) As String
'Returns the nth word of a text
    Dim nWO_i As Long, nWO_Tmp As Long
    
    For nWO_i = 1 To n - 1
        nWO_Tmp = InStr(TextToSeperate, " ")
        If nWO_Tmp > 0 Then
            TextToSeperate = Right(TextToSeperate, Len(TextToSeperate) - nWO_Tmp)
        End If
    Next nWO_i
    nWO_Tmp = InStr(TextToSeperate, " ")
    If nWO_Tmp > 0 Then 'If this is not the last word
        TextToSeperate = Left(TextToSeperate, nWO_Tmp - 1)
    End If
    nthWordOf = TextToSeperate
End Function

Function MaxWidthWordInColumn(ColId) As Single
    Dim MWW_i As Long, MWW_j As Long
    Dim MWW_Tmp As Single, MWW_Max As Single
    Dim MWW_sTmp As String
    
    MWW_Max = 0: MWW_Tmp = 0
    For MWW_i = 0 To FlexGrid2.Rows - 1
        MWW_sTmp = FlexGrid2.TextMatrix(MWW_i, ColId)
        For MWW_j = 1 To WordsIn(MWW_sTmp)
            MWW_Tmp = Printer.TextWidth(nthWordOf(MWW_sTmp, MWW_j))
            If MWW_Tmp > MWW_Max Then MWW_Max = MWW_Tmp
        Next MWW_j
    Next MWW_i
    MaxWidthWordInColumn = MWW_Max
End Function

Function ParseTextToWidth(ByVal TextToParse As String, ByVal WidthInmm As Single, ByVal RowId As Long, ByVal ColId As Long) As Ox_ParseResult
'Returns the new width, height of the cell and the word groups of text as parseresult type
    Dim PTTW_i As Long, PTTW_j As Long, PTTW_sTmp As String
    Dim PTTW_GroupTmp As String, PTTW_Tmp As Long
    Dim PTTW_AvailableWidth As Long, PTTW_WordCount As Long
    Dim PTTW_TmpHeight As Long, PTTW_TotalHeight As Long
    Dim PTTW_TmpWidth As Long, PTTW_MaxWidth As Single
    Dim PTTW_GroupCount As Long, PTTW_LastNumOfWordUsed
    Dim tmp As Single
    
'if a group is above the width limit, then resize the cell and re-start parsing
ReStartParsing:
    PTTW_MaxWidth = WidthInmm
    PTTW_TotalHeight = 0 'Values will be added
    PTTW_GroupCount = 0
    PTTW_WordCount = WordsIn(TextToParse)
    For PTTW_i = 1 To PTTW_WordCount
        
        PTTW_GroupTmp = nthWordOf(TextToParse, PTTW_i) 'Take the nth word as a group
        PTTW_LastNumOfWordUsed = PTTW_i

CheckGroupWidth:
        tmp = Printer.TextWidth(PTTW_GroupTmp)
        If Printer.TextWidth(PTTW_GroupTmp) > WidthInmm Then 'Check the width
            If WordsIn(PTTW_GroupTmp) > 1 Then 'It is a multi word group
                PTTW_GroupTmp = CutLastWord(PTTW_GroupTmp) 'OxF. Cuts the last word of a string
            End If

AddGroup:
            PTTW_GroupCount = PTTW_GroupCount + 1
            ReDim Preserve ParseTextToWidth.Groups(PTTW_GroupCount)
            ParseTextToWidth.Groups(PTTW_GroupCount) = PTTW_GroupTmp
            If PTTW_MaxWidth < Printer.TextWidth(PTTW_GroupTmp) Then PTTW_MaxWidth = Printer.TextWidth(PTTW_GroupTmp)
            If WordsIn(TextToParse) > 1 Then
                PTTW_i = PTTW_LastNumOfWordUsed - 1 'Set the counter to the last word
            End If
            'Add to height
            PTTW_TotalHeight = PTTW_TotalHeight + Printer.TextHeight(PTTW_GroupTmp) + 1 '1 is for space between groups
            
            If PTTW_MaxWidth > WidthInmm Then 'If the width is changed, start again
                WidthInmm = PTTW_MaxWidth + 1 '2 for left, 2 for right
                GoTo ReStartParsing
            End If
        
        Else 'If it is in limits, add the next word to the group
            If PTTW_LastNumOfWordUsed < PTTW_WordCount Then
                PTTW_GroupTmp = PTTW_GroupTmp & " " & nthWordOf(TextToParse, PTTW_LastNumOfWordUsed + 1)
                PTTW_LastNumOfWordUsed = PTTW_LastNumOfWordUsed + 1
                GoTo CheckGroupWidth 'Re-Check the group width
            Else
                PTTW_LastNumOfWordUsed = PTTW_LastNumOfWordUsed + 1 'For avoiding duplication
                GoTo AddGroup
            End If
        End If
    
    Next PTTW_i
    
    'Now we have to count the new height and width
    '2 mm between groups, 2 mm at top&bottom, left&right
    
    If PTTW_TotalHeight > TwipsTomm(FlexGrid2.CellHeight) Then
        FlexGrid2.RowHeight(RowId) = FlexGrid2.RowHeight(RowId) + mmToTwips(PTTW_TotalHeight) - FlexGrid2.CellHeight
        'Add the difference to the first row's height
    End If
    If PTTW_MaxWidth > TwipsTomm(FlexGrid2.CellWidth) Then
        FlexGrid2.ColWidth(ColId) = FlexGrid2.ColWidth(ColId) + mmToTwips(PTTW_MaxWidth + 8) - FlexGrid2.CellWidth  '4 for left, 4 for right
        'Add the difference to the first col's width
    End If
    
    With ParseTextToWidth
        .GroupCount = PTTW_GroupCount
        .Alignment = FlexGrid2.CellAlignment
        .BackColor = FlexGrid2.CellBackColor
    End With
    
End Function

Function TotalWidthOfFlexGrid()
    Dim TW_i As Long, TW_j As Long
    
    For TW_i = 0 To FlexGrid2.Cols - 1
        TotalWidthOfFlexGrid = TotalWidthOfFlexGrid + FlexGrid2.ColWidth(TW_i)
    Next TW_i
    TW_j = TotalWidthOfFlexGrid
    TotalWidthOfFlexGrid = TwipsTomm(TW_j)
End Function

Function TotalHeight(ByVal RowId As Long, ByVal ColId As Long) As Single
'Returns the total height of text groups in a cell
    Dim TH_i As Long
    
    TotalHeight = 0.2
    With CellInfos(RowId, ColId)
    For TH_i = 1 To .GroupCount
        TotalHeight = TotalHeight + Printer.TextHeight(.Groups(TH_i)) + 0.2
    Next TH_i
    End With
End Function

Sub DrawCellToPrinter(ByVal x As Single, ByVal y As Single, ByVal RowId As Long, ByVal ColId As Long)
'Draws the current cell to printer, and prints the text in it
    Dim DCTP_Alignment As String
    Dim DCTP_Textx As Single, DCTP_Texty As Single
    Dim DCTP_i As Long, DCTP_j As Long, DCTP_TextTmp As String
    Dim DCTP_TextStarty As Single, DCTP_TextTotalHeight As Single
    
    FlexGrid2.Row = RowId
    FlexGrid2.Col = ColId
    
    DCTP_Alignment = CountAlignment(FlexGrid2.CellAlignment)
    
    If TableBorders = True Then 'If user wants border lines or not
        If PrintToPicBox = True Then
            'NONEED
            'Dim tmp As Long
            'tmp = 150 + Int(Rnd * 274)
            FrmPrintFlex.Picture1.Line (x, y)-(x + TwipsTomm(FlexGrid2.CellWidth), y + TwipsTomm(FlexGrid2.CellHeight)), RGB(0, 0, 0), B
        Else
            'tmp = 150 + Int(Rnd * 274)
            Printer.Line (x, y)-(x + TwipsTomm(FlexGrid2.CellWidth), y + TwipsTomm(FlexGrid2.CellHeight)), RGB(0, 0, 0), B
        End If
    End If

'Print the text
    DCTP_TextTotalHeight = TotalHeight(RowId, ColId)
    DCTP_TextStarty = y + 0.5 + (TwipsTomm(FlexGrid2.CellHeight) - DCTP_TextTotalHeight) / 2
    With CellInfos(RowId, ColId)
    For DCTP_i = 1 To .GroupCount
        DCTP_TextTmp = .Groups(DCTP_i)
        Select Case CountAlignment(.Alignment)
            Case "Left"
                DCTP_Textx = x + 0.8
            Case "Right"
                DCTP_Textx = x - 0.8 + TwipsTomm(FlexGrid2.CellWidth) - Printer.TextWidth(DCTP_TextTmp)
            Case "Center"
                DCTP_Textx = x + (TwipsTomm(FlexGrid2.CellWidth) - Printer.TextWidth(DCTP_TextTmp)) / 2
        End Select
        DCTP_Texty = Printer.TextHeight(DCTP_TextTmp)
        DCTP_Texty = DCTP_TextStarty + (DCTP_i - 1) * (Printer.TextHeight(DCTP_TextTmp) + 0.2)
        If PrintToPicBox = True Then
            FrmPrintFlex.Picture1.CurrentX = DCTP_Textx
            FrmPrintFlex.Picture1.CurrentY = DCTP_Texty
            FrmPrintFlex.Picture1.Print DCTP_TextTmp
        Else
            Printer.CurrentX = DCTP_Textx
            Printer.CurrentY = DCTP_Texty
            Printer.Print DCTP_TextTmp
        End If
    Next DCTP_i
    End With
End Sub

Function NoOfPages() As Long
'Returns the number of pages needed to print the flex grid
    Dim NOP_i As Long, NOP_j As Long
    Dim NOP_StartRow As Long, NOP_nRow As Long
    Dim NOP_n As Long, NOP_Starty As Single
    Dim NOP_StartRowTmp As Long
    
    NOP_n = 0
    NOP_StartRow = nHeaderRows

NextPage:
    NOP_nRow = 0
    NOP_StartRowTmp = NOP_StartRow
    If NOP_StartRow = FlexGrid.Rows - nFooterRows Then 'Pages ended, Full fix!
        NoOfPages = NOP_n
        If NoOfPages = 0 Then NoOfPages = 1 'One page only..
        Exit Function
    End If
    NOP_n = NOP_n + 1
    'NOP_nRow = 0
    NOP_Starty = CellInfos(NOP_StartRow, 0).Defaulty
    
    Do While CellInfos(NOP_StartRow, 0).Defaulty - NOP_Starty <= PrHWithoutHAndF
        NOP_nRow = NOP_nRow + 1
        NOP_StartRow = NOP_StartRow + 1
        If NOP_StartRow = FlexGrid.Rows - nFooterRows Then 'Pages ended
            NoOfPages = NOP_n
            ReDim Preserve PageInfos(NOP_n)
            PageInfos(NOP_n).nRows = NOP_nRow
            PageInfos(NOP_n).StartRow = NOP_StartRowTmp
            PageInfos(NOP_n).PageHeightInmm = CellInfos(NOP_StartRow - 1, 0).Defaulty + TwipsTomm(FlexGrid2.RowHeight(NOP_StartRow - 1)) - NOP_Starty
            Exit Function
        End If
    'If NOP_StartRow = FlexGrid.Rows - nFooterRows Then
    '    Exit Do
    'End If
    Loop
    NOP_StartRow = NOP_StartRow - 2 '1 for adding at start, 1 for avoid overflow
    'NOP_n = NOP_n + 1
    ReDim Preserve PageInfos(NOP_n)
    PageInfos(NOP_n).nRows = NOP_nRow - 2
    PageInfos(NOP_n).StartRow = NOP_StartRowTmp
    PageInfos(NOP_n).PageHeightInmm = CellInfos(NOP_StartRow, 0).Defaulty - NOP_Starty
    GoTo NextPage

End Function

Sub PrintHeader(ByVal PageId As Long)
'Prints the header
    Dim PH_i As Long, PH_j As Long
    
    'Print the table text first;
    If PrintToPicBox = True Then
        FrmPrintFlex.Picture1.CurrentX = Margins.LeftInmm
        FrmPrintFlex.Picture1.CurrentY = Margins.TopInmm
        FrmPrintFlex.Picture1.Print "Table created by WMRP2000, page " & CStr(PageId) & "/" & CStr(NumOfPages) & ",  " & CStr(Day(Now)) & "." & CStr(Month(Now)) & "." & CStr(Year(Now))
    Else
        Printer.CurrentX = Margins.LeftInmm
        Printer.CurrentY = Margins.TopInmm
        Printer.Print "Table created by WMRP2000, page " & CStr(PageId) & "/" & CStr(NumOfPages) & ",  " & CStr(Day(Now)) & "." & CStr(Month(Now)) & "." & CStr(Year(Now))
    End If
    If nHeaderRows = 0 Then Exit Sub
    For PH_i = 0 To nHeaderRows - 1
        For PH_j = 0 To FlexGrid.Cols - 1
            If IsCellHorizontalMerged(PH_i, PH_j) = False And IsCellVerticalMerged(PH_i, PH_j) = False Then
                'Cell is not merged in any directions
                DrawCellToPrinter CellInfos(PH_i, PH_j).x + Margins.LeftInmm, CellInfos(PH_i, PH_j).y + TableTextHeight * 2 + Margins.TopInmm, PH_i, PH_j
            End If
        Next PH_j
    Next PH_i
End Sub

Sub PrintPage(ByVal PageId As Long)
'Prints the page with the page Id
    Dim PP_i As Long, PP_j As Long, PP_k As Long
    Dim PP_CellInfo As Ox_ParseResult
    
    PrintHeader (PageId) 'OxF. Print the header for each page
    
    For PP_i = PageInfos(PageId).StartRow To PageInfos(PageId).StartRow + PageInfos(PageId).nRows - 1
        For PP_j = 0 To FlexGrid.Cols - 1
            If IsCellHorizontalMerged(PP_i, PP_j) = False And IsCellVerticalMerged(PP_i, PP_j) = False Then
                'Cell is not merged in any directions
                DrawCellToPrinter CellInfos(PP_i, PP_j).x + Margins.LeftInmm, CellInfos(PP_i, PP_j).y + TableTextHeight * 2 + Margins.TopInmm - (CellInfos(PageInfos(PageId).StartRow, 0).Defaulty - CellInfos(nHeaderRows, 0).Defaulty), PP_i, PP_j
            End If
        Next PP_j
    Next PP_i
    CurrentPageFooterStartPosition = CellInfos(PP_i - 1, 0).Defaulty + TwipsTomm(FlexGrid2.RowHeight(PP_i - 1)) - (CellInfos(PageInfos(PageId).StartRow, 0).Defaulty - CellInfos(nHeaderRows, 0).Defaulty)
    PrintFooter (PageId) 'OxF. Print the footer for each page
End Sub

Sub PrintFooter(ByVal PageId As Long)
'Prints the footer at the necessary position for that page
    Dim PF_i As Long, PF_j As Long, PF_Tmp As Long
    Dim PF_cellinfo As Ox_ParseResult
    
    If nFooterRows = 0 Then Exit Sub
    PF_Tmp = FlexGrid.Rows - nFooterRows 'Footer start row
    For PF_i = PF_Tmp To FlexGrid.Rows - 1
        For PF_j = 0 To FlexGrid.Cols - 1
            If IsCellHorizontalMerged(PF_i, PF_j) = False And IsCellVerticalMerged(PF_i, PF_j) = False Then
                'Cell is not merged in any directions
                DrawCellToPrinter CellInfos(PF_i, PF_j).x + Margins.LeftInmm, CellInfos(PF_i, PF_j).y + TableTextHeight * 2 + Margins.TopInmm - (CellInfos(FlexGrid2.Rows - nFooterRows, 0).Defaulty - CurrentPageFooterStartPosition), PF_i, PF_j
            End If
        Next PF_j
    Next PF_i
    'Print the page number
    If PrintPageNo = True Then
        If PrintToPicBox = True Then
            FrmPrintFlex.Picture1.CurrentX = TwipsTomm(FrmPrintFlex.Picture1.Width) - Margins.RightInmm
            FrmPrintFlex.Picture1.CurrentY = TwipsTomm(FrmPrintFlex.Picture1.Height) - Margins.BottomInmm
            FrmPrintFlex.Picture1.Print CStr(PageId)
        Else
            Printer.CurrentX = TwipsTomm(Printer.Width) - Margins.RightInmm
            Printer.CurrentY = TwipsTomm(Printer.Height) - Margins.BottomInmm
            Printer.Print CStr(PageId)
        End If
    End If
End Sub

Function HeaderHeight() As Long
'Returns the height of header
    Dim HH_i As Long
    
    If nHeaderRows = 0 Then Exit Function
    For HH_i = 0 To FlexGrid2.Cols - 1
        If IsCellVerticalMerged(nHeaderRows, HH_i) = False Then
            HeaderHeight = CellInfos(nHeaderRows, HH_i).y
            Exit Function
        End If
    Next HH_i
    HeaderHeight = CellInfos(nHeaderRows).y
End Function

Function FooterHeight() As Long
'Returns the height of the footer
    Dim FH_i As Long
    
    If nFooterRows = 0 Then Exit Function
    For FH_i = 0 To FlexGrid2.Cols - 1
        If IsCellVerticalMerged(FlexGrid2.Rows - nFooterRows, FH_i) = False Then
            FooterHeight = FGHInmm - CellInfos(FlexGrid2.Rows - nFooterRows, FH_i).y
            Exit Function
        End If
    Next FH_i
    FooterHeight = FGHInmm - CellInfos(FlexGrid2.Rows - nFooterRows).y
End Function

Function IsCellVerticalMerged(ByVal RowId As Long, ByVal ColId As Long) As Boolean
'Returns if the cell is a member of vertical merged cells
    If RowId = 0 Then Exit Function
    If FlexGrid.MergeCol(ColId) = True Then
        If FlexGrid.TextMatrix(RowId, ColId) = FlexGrid.TextMatrix(RowId - 1, ColId) Then
            IsCellVerticalMerged = True
        End If
    End If
End Function

Function IsCellHorizontalMerged(ByVal RowId As Long, ByVal ColId As Long) As Boolean
'Returns if the cell is a member of horizontal merged cells
    If ColId = 0 Then Exit Function
    If FlexGrid.MergeRow(RowId) = True Then
        If FlexGrid.TextMatrix(RowId, ColId) = FlexGrid.TextMatrix(RowId, ColId - 1) Then
            IsCellHorizontalMerged = True
        End If
    End If
End Function
    
Function TwipsTomm(Tw As Long) As Single
'Returns the mm compensation of the specified amount of twips
'567 twips is appr. 1 cm, 56.7 twips is appr. 1 mm
    TwipsTomm = Tw / 56.7
End Function

Function mmToTwips(mm As Long) As Single
'Returns the mm compensation of the specified amount of twips
'567 twips is appr. 1 cm, 56.7 twips is appr. 1 mm
    mmToTwips = mm * 56.7
End Function

Sub FillFormats()
'Fills the format variables for each col/row of flex grid
    Dim FF_i As Long, FF_j As Long
    
    For FF_i = 0 To FlexGrid.Rows - 1
        FlexGrid.Row = FF_i
        For FF_j = 0 To FlexGrid.Cols - 1
            FlexGrid.Col = FF_j
            CellInfos(FF_i, FF_j).NewHeightInmm = TwipsTomm(FlexGrid.CellHeight)
            CellInfos(FF_i, FF_j).NewWidthInmm = TwipsTomm(FlexGrid.CellWidth)
        Next FF_j
    Next FF_i
End Sub

Function FirstCellOfHorizontalGroup(ByVal RowId As Long, ByVal ColId As Long)
    'Returns the first cell of an horizontally merged group
    Dim FCOHG_i As Long, FCOHG_j As Long
        
        For FCOHG_i = ColId To 0 Step -1
            If IsCellHorizontalMerged(RowId, FCOHG_i) = False Then
                FirstCellOfHorizontalGroup = FCOHG_i
                Exit Function
            End If
        Next FCOHG_i
End Function

Function FirstCellOfVerticalGroup(ByVal RowId As Long, ByVal ColId As Long)
    'Returns the first cell of an horizontally merged group
    Dim FCOVG_i As Long
        
        For FCOVG_i = RowId To 0 Step -1
            If IsCellVerticalMerged(FCOVG_i, ColId) = False Then
                FirstCellOfVerticalGroup = FCOVG_i
                Exit Function
            End If
        Next FCOVG_i
End Function

Function ExploreUserFormat(FStr As String)
'Explores the user format string and updates the specified col/row format variables
    Dim EUF_i As Long, EUF_j As Long, EUF_Tmp As Long, EUF_sTmp As String
    Dim EUF_ColOrRow As String * 1, EUF_Num As Long, EUF_Value As Long
    Dim EUF_ColWidth As Long, EUF_RowHeight As Long, EUF_CellToChange As Ox_Cell
    
    EUF_Tmp = Len(FStr)
    For EUF_i = 1 To EUF_Tmp Step 8 'For each formatted col/row
        EUF_sTmp = Mid(FStr, EUF_i, 8) 'Get the next part
        EUF_ColOrRow = Left(EUF_sTmp, 1)
        EUF_Num = Val(Mid(EUF_sTmp, 2, 3))
        EUF_Value = Val(Mid(EUF_sTmp, 5, 4))
        Select Case EUF_ColOrRow
        
        Case "c" 'If a col will be resized
            FlexGrid2.ColWidth(EUF_Num) = mmToTwips(EUF_Value)
        
        Case "r"
            FlexGrid2.RowHeight(EUF_Num) = mmToTwips(EUF_Value)
        
        End Select
    Next EUF_i
End Function


Sub CalculateStartingPositions()
'Calculates the starting positions by using the width and height values
    Dim i As Long, j As Long, k As Long, tmp As Long
    Dim Totalx As Long, Totaly() As Long
    Dim MainStartx As Single, MainStarty As Single
    
    MainStartx = 0
    MainStarty = 0
    
    ReDim Totaly(FlexGrid2.Cols)
    For i = 0 To FlexGrid2.Rows - 1
        Totalx = 0
        FlexGrid2.Row = i
        
        For j = 0 To FlexGrid2.Cols - 1
            FlexGrid2.Col = j
            
            For k = 0 To j - 1
                CellInfos(i, j).Defaultx = CellInfos(i, j).Defaultx + TwipsTomm(FlexGrid2.ColWidth(k))
            Next k
            
            For k = 0 To i - 1
                CellInfos(i, j).Defaulty = CellInfos(i, j).Defaulty + TwipsTomm(FlexGrid2.RowHeight(k))
            Next k
            
            If IsCellVerticalMerged(i, j) = True Then
                CellInfos(i, j).y = CellInfos(i - 1, j).y
            Else
                CellInfos(i, j).y = TwipsTomm(Totaly(j))
                Totaly(j) = Totaly(j) + FlexGrid2.RowHeight(i)
                For k = i + 1 To FlexGrid2.Rows - 1
                    If IsCellVerticalMerged(k, j) = True Then
                        Totaly(j) = Totaly(j) + FlexGrid2.RowHeight(k)
                    Else
                        Exit For
                    End If
                Next k
            End If
            
            If IsCellHorizontalMerged(i, j) = True Then
                CellInfos(i, j).x = CellInfos(i, j - 1).x
            Else
                CellInfos(i, j).x = TwipsTomm(Totalx)
                Totalx = Totalx + FlexGrid2.ColWidth(j)
                For k = j + 1 To FlexGrid2.Cols - 1
                    If IsCellHorizontalMerged(i, k) = True Then
                        Totalx = Totalx + FlexGrid2.ColWidth(k)
                    Else
                        Exit For
                    End If
                Next k
            End If
            
        Next j
    Next i
    
    FGHInmm = 0: FGWInmm = 0
    For k = 0 To FlexGrid2.Rows - 1
        FGHInmm = FGHInmm + TwipsTomm(FlexGrid2.RowHeight(k))
    Next k
    
    For k = 0 To FlexGrid2.Cols - 1
        FGWInmm = FGWInmm + TwipsTomm(FlexGrid2.ColWidth(k))
    Next k

End Sub

Sub CopyFlexGrid(Fg1 As MSFlexGrid, Fg2 As MSFlexGrid)
'Copies a flex grid to another
    Dim CFG_i As Long, CFG_j As Long, CFG_Color As Long
    
    Fg2.Rows = Fg1.Rows
    Fg2.Cols = Fg1.Cols
    Fg2.FixedCols = Fg1.FixedCols
    Fg2.FixedRows = Fg1.FixedRows
    Fg2.Width = Fg1.Width
    Fg2.Height = Fg1.Height
    Fg2.MergeCells = Fg1.MergeCells
    
    For CFG_i = 0 To Fg1.Rows - 1
        Fg1.Row = CFG_i
        Fg2.Row = CFG_i
        Fg2.RowHeight(CFG_i) = Fg1.RowHeight(CFG_i) 'xxx
        Fg2.MergeRow(CFG_i) = Fg1.MergeRow(CFG_i)
        For CFG_j = 0 To Fg1.Cols - 1
            Fg2.MergeCol(CFG_j) = Fg1.MergeCol(CFG_j)
            Fg1.Col = CFG_j
            Fg2.Col = CFG_j
            Fg2.ColWidth(CFG_j) = Fg1.ColWidth(CFG_j)
            Fg2.TextMatrix(CFG_i, CFG_j) = Fg1.TextMatrix(CFG_i, CFG_j)
            Fg2.CellBackColor = Fg1.CellBackColor
            Fg2.CellAlignment = Fg1.CellAlignment
            Fg2.CellFontBold = Fg1.CellFontBold
            Fg2.CellFontItalic = Fg1.CellFontItalic
            Fg2.CellFontName = Fg1.CellFontName
            Fg2.CellFontSize = Fg1.CellFontSize
        Next CFG_j
    Next CFG_i
End Sub

'The main routine for calling from an outside procedure..

Sub QuickPrintFlexGrid(ByVal FgContainer As Form, ByVal Fg1 As MSFlexGrid, ByVal Fg2 As MSFlexGrid, ByVal CD As CommonDialog, ByVal nHeader As Long, ByVal nFooter As Long, ByVal FStr As String, TrnMargins As Ox_Margins, ByVal PageNumbers As Boolean, ByVal TBorders As Boolean)
'Prints the flex grid in a B&W line style format

'User format string format is : "cnnnxxxx" or "rnnnxxxx", c for col, r for row,
'nnn for row/col number, xxxx for new widht/height in mm's (Fixed length: 8 bytes)
    
'OxF at the comments means the statement calls a function or a sub routine in this module

    'Let's print!
    PrintToPicBox = True 'Set this to check it on a pic box
    PrintPageNo = PageNumbers
    TableBorders = TBorders
    For i = 1 To NumOfPages
        FrmPrintFlex.Picture1.Height = Printer.Height
        FrmPrintFlex.Picture1.Width = Printer.Width
        FrmPrintFlex.Picture1.Picture = Nothing
        PrintPage (i)
        Clipboard.Clear
        Clipboard.SetData FrmPrintFlex.Picture1.Image
        If i = NumOfPages Then Exit For
        
'        MsgBox ""
        Printer.NewPage
    Next i
    'Printer.KillDoc
    Printer.EndDoc
    
    FgContainer.ScaleMode = TmpScaleMode 'Set the current scale mode back

End Sub

Function GeneratePreview(ByVal FgContainer As Form, ByVal Fg1 As MSFlexGrid, ByVal Fg2 As MSFlexGrid, ByVal CD As CommonDialog, ByVal nHeader As Long, ByVal nFooter As Long, ByVal FStr As String, TrnMargins As Ox_Margins, TrnFontSize As Single) As Ox_TableProp
    
    Dim tmp As Long
    
    'CD.ShowPrinter 'Show printer dialog box and set the font size
    Printer.ScaleMode = vbMillimeters
    Printer.DrawWidth = 0.2 * 56.7 / Printer.TwipsPerPixelX '.2 mm, let user enter the value!
    Printer.FontName = "arial tur"
    Printer.FontSize = TrnFontSize
    Printer.ForeColor = RGB(0, 0, 0)
    FrmPrintFlex.Picture1.ScaleMode = vbMillimeters
    FrmPrintFlex.Picture1.DrawWidth = 0.2 * 56.7 / Screen.TwipsPerPixelX '.2 mm, let user enter the value!
    FrmPrintFlex.Picture1.FontName = "Arial Tur"
    FrmPrintFlex.Picture1.FontSize = TrnFontSize
    FrmPrintFlex.Picture1.ForeColor = RGB(0, 0, 0)
    
    Margins.BottomInmm = TrnMargins.BottomInmm
    Margins.LeftInmm = TrnMargins.LeftInmm
    Margins.RightInmm = TrnMargins.RightInmm
    Margins.TopInmm = TrnMargins.TopInmm
    TableTextHeight = Printer.TextHeight("Table cre")
    nHeaderRows = nHeader 'Header and footer number of rows
    nFooterRows = nFooter
    
    CopyFlexGrid Fg1, Fg2 'copies fg1 to fg2
    
    Set FlexGrid = Fg1 'Sets the flex grid for use in other functions/sub routines
    Set FlexGrid2 = Fg2
    ReDim CellInfos(FlexGrid.Rows - 1, FlexGrid.Cols - 1)
    ReDim MaxWidthWordInColumns(FlexGrid.Cols - 1)
    
    TmpScaleMode = FgContainer.ScaleMode 'Store the current scalemode for the flex grid's container
    FgContainer.ScaleMode = vbTwips 'Set to twips
    
    PrWInmm = TwipsTomm(Printer.Width) 'Getting page height & width
    PrHInmm = TwipsTomm(Printer.Height)
    
    'Fg.Visible = False 'Set to unvisible to avoid cell-selecting effects
    
    TmpRow = Fg1.Row 'Store the current cell selection
    TmpCol = Fg1.Col
    
    'NONEED
    'FillFormats 'OxF. Fills the cell format variables by using the flex grid
    'A big amount of these values will be changed after parsing
    
    ExploreUserFormat (FStr) ' OxF. Explores the user format and store the values
    
    'Calculating new height and width values
    'First step is important too. Calculate the longest word
    'In the column..
    For j = 0 To FlexGrid.Cols - 1
        MaxWidthWordInColumns(j) = MaxWidthWordInColumn(j)
    Next j
    
    For i = 0 To FlexGrid.Rows - 1
        FlexGrid2.Row = i
        For j = 0 To FlexGrid.Cols - 1
            FlexGrid2.Col = j
            
            If IsCellVerticalMerged(i, j) = False And IsCellHorizontalMerged(i, j) = False Then
                CellInfos(i, j) = ParseTextToWidth(FlexGrid2.TextMatrix(i, j), MaxWidthWordInColumns(j), i, j)
                'OxF. Update the cell info by parsing the text to the width of cell
            End If
                
        Next j
    Next i
    
    'NONEED
    'CorrectCellInfos 'OxF. Sets the width/height to the max value in the col/row
    
    CalculateStartingPositions 'OxF. Calculates the starting positions by using the width and height values
        
    HeaderHeightInmm = HeaderHeight
    FooterHeightInmm = FooterHeight
    
    PrHWithoutHAndF = PrHInmm - HeaderHeightInmm - FooterHeightInmm - Margins.BottomInmm - Margins.TopInmm - TableTextHeight * 2
    
    NumOfPages = NoOfPages 'OxF. Calculates the number of pages needed to print all of the flex grid

    With GeneratePreview
    .NumOfPages = NumOfPages
    .TableWidth = TotalWidthOfFlexGrid
    ReDim .TableHeights(NumOfPages)
    For tmp = 1 To NumOfPages
        .TableHeights(tmp) = HeaderHeightInmm + PageInfos(tmp).PageHeightInmm + FooterHeightInmm
    Next tmp
    End With
End Function
