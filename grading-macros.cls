VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''Excel VBA Grading Sheet Macros
''MIT License

''Copyright (c) 2022 Joseph Timothy Foley <foley AT RU dot IS>
''
''Permission is hereby granted, free of charge, to any person obtaining a copy
''of this software and associated documentation files (the "Software"), to deal
''in the Software without restriction, including without limitation the rights
''to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
''copies of the Software, and to permit persons to whom the Software is
''furnished to do so, subject to the following conditions:
''
''The above copyright notice and this permission notice shall be included in all
''copies or substantial portions of the Software.
''
''THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
''OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
''SOFTWARE.

Option Explicit
Public strNamelocation As String
Public strNameRange As String
Public strReviewerrange As String
Public strTemplatename As String
Public WBdata As Workbook
Public boolGlobalsLoaded As Boolean

Sub InitGlobals()                                'Called by Sheet_Open to setup when sheet is opened
    Dim ws As Worksheet
    Dim datasource As String
    Set ws = Application.ThisWorkbook.Sheets(1)
    'If boolGlobalsLoaded <> True Then
    With ws
        strNamelocation = .Range("B3").Value     'Usually B2 or E2
        strNameRange = .Range("B4").Value        'Usually B2:D2
        strReviewerrange = .Range("B5").Value    'Usually B3:E3
        strTemplatename = .Range("B6").Value     'Usually template
        ' TODO:  Automatic detection of sheet name if not set
        datasource = .Range("B7").Value
        Select Case datasource
        Case "ThisWorkbook"
            Set WBdata = Application.ThisWorkbook 'for current sheet processing.  Needs tuning.
        Case "", "ActiveWorkbook"
            Set WBdata = ActiveWorkbook
        Case Else
            Set WBdata = Workbooks(datasource)   'assume sheet is already opened
            'TODO:  file if not already opened
        End Select
    End With
    boolGlobalsLoaded = True
    'End If
End Sub

Sub PrepareAndPrint()
    Call InitGlobals
    Call PrepareSheetsForDistribution
    Call ExportToPDFs
End Sub

Sub PrepareSheetsForDistribution()
    ' fix inconsistencies and sterilize
    Call InitGlobals
    Dim ws As Worksheet
    Call SyncTabToCellname
    Call SortWorksheetsTabs
    For Each ws In WBdata.Worksheets
        ' Skip processing the sheets that start as template
        If InStr(1, ws.name, strTemplatename) <> 1 Then
            '' sanitize reviewer except template for later checking
            ws.Range(strReviewerrange).Value = ""
        End If
        ' sane values for PDF output
        With ws.PageSetup
            .PaperSize = xlPaperA4
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False              'width sets number of pages
        End With

    Next ws
End Sub

Sub ExportToPDFs()
    ' Export all sheets by name into pdf
    Call InitGlobals
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    For Each ws In WBdata.Worksheets
        ' Skip processing the sheets that start as template
        If InStr(1, ws.name, strTemplatename) <> 1 Then
            Call ExportSheetToPDF(ws)
        End If
    Next ws
    Application.ScreenUpdating = True

End Sub

Function GetFirstInRow(SearchRange As String, ws As Worksheet) As Variant
    Call InitGlobals
    GetFirstInRow = ws.Range(SearchRange).Find(What:="*", _
                                               LookAt:=xlPart, LookIn:=xlValues, SearchOrder:=xlByColumns, _
                                               SearchDirection:=xlNext, MatchCase:=False)
End Function

Sub SyncTabToCellname()
    Call InitGlobals
    ' Fix inconsistent numbers, names, and formatting
    Dim ws As Worksheet
    Dim teamindex As Integer
    Dim teamstring As String
    teamindex = 0
    For Each ws In WBdata.Worksheets
        ' Skip processing the sheets that start as template
        If InStr(1, ws.name, strTemplatename) <> 1 Then
            'teamindex = ws.Range(strNamelocation).Value
            teamindex = GetFirstInRow(strNameRange, ws)
            'teamstring = "team" & Format(teamindex, "00") ''decided to keep it simpler
            teamstring = Format(teamindex, "00")
            ws.name = teamstring
        End If
    Next ws
End Sub

Sub SortWorksheetsTabs()
    Call InitGlobals
    'Modified from https://trumpexcel.com/sort-worksheets/
    Application.ScreenUpdating = False
    Dim ShCount As Integer, i As Integer, j As Integer
    ShCount = WBdata.Sheets.Count
    Dim ws As Worksheet, templateendindex As Integer
    templateendindex = 0
    For Each ws In WBdata.Worksheets             'count templates at beginning
        ' Skip processing the sheets that start as template
        If InStr(1, ws.name, strTemplatename) = 1 Then
            templateendindex = templateendindex + 1
        End If
        Debug.Print templateendindex
    Next ws
    
    For i = 1 + templateendindex To ShCount - 1
        For j = i + 1 To ShCount
            If UCase(WBdata.Sheets(j).name) < UCase(WBdata.Sheets(i).name) Then
                WBdata.Sheets(j).Move before:=WBdata.Sheets(i)
            End If
        Next j
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub MakeSheetInstance(CreationIndex As Integer, name As String)
    'Copy the template sheet to another location
    'Modes:  'query' (default), '
    Call InitGlobals
    Dim templatews As Worksheet
    Dim myCellSetValue As Range
    Set templatews = WBdata.Sheets(1)
    ' Adjust above if you need to change the template but should be at the beginning
    ' TODO: search for first sheet named template
    With WBdata
        templatews.Copy After:=.Sheets(CreationIndex)
        .ActiveSheet.name = name
        Set myCellSetValue = .ActiveSheet.Range(strNamelocation)
        myCellSetValue.Value = name
    End With
End Sub

Sub MakeSheetInstanceEnd()
    Call InitGlobals
    Dim name As String
    name = InputBox("Sheet Name")
    Call MakeSheetInstance(WBdata.Sheets.Count, name)
End Sub

Sub MakeSheetInstanceNext()
    Call InitGlobals
    Dim name As String
    name = InputBox("Sheet Name")
    Call MakeSheetInstance(WBdata.ActiveSheet.Index)
End Sub

Function ExportSheetToPDF(ws As Worksheet)
    Call InitGlobals
    Dim pdfname As String, ThisFileName As String, FileNameArray() As String
    Dim BaseFileName As String, nm As String
    
    ThisFileName = WBdata.name                   '.FullName gives everything, .Path gives path
    FileNameArray = Split(ThisFileName, ".")
    BaseFileName = FileNameArray(0)
    
    nm = ws.name
    pdfname = Application.ActiveWorkbook.Path & "\" & BaseFileName & _
              "-" & nm & ".pdf"
    
    'Debug.Print pdfname
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           Filename:=pdfname, _
                           Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, OpenAfterPublish:=False

End Function

Sub ExportCurrentSheetToPDF()
    Call InitGlobals
    Call ExportSheetToPDF(WBdata.ActiveSheet)
End Sub


