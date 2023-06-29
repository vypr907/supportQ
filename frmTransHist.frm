VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTransHist 
   Caption         =   "Transaction History"
   ClientHeight    =   8880.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13950
   OleObjectBlob   =   "frmTransHist.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTransHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnPrint_Click()
'sub to print transaction log to PDF
    Dim fileName As Variant
    Dim log As Worksheet
    
    Set log = ActiveWorkbook.Worksheets("transactionLog")
    
    'dialog box to get location and name for pdf file
    fileName = Application.GetSaveAsFilename( _
        InitialFileName:="Transaction Log", _
        FileFilter:="PDF, *.pdf", _
        title:="Save as PDF")
        
    If fileName <> False Then 'if user doesn't hit cancel
    
        With log.PageSetup
            .CenterHeader = "Transaction Log"
            .Orientation = xlPortrait
            .PrintTitleRows = log.Rows(1).Address
            .Zoom = False
            .FitToPagesTall = False
            .FitToPagesWide = 1
        End With
        
        log.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=fileName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    'start userform centered inside Excel screen-----------
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    '------------------------------------------------------


    'loads transactionLog into listbox
    'Dim cell As Range
    Dim lastRow As Long
    lastRow = Worksheets("transactionLog").Cells(Rows.Count, 2).End(xlUp).Row
    
    'format columns
    ListBox1.ColumnCount = 5
    ListBox1.ColumnWidths = "100,50,50,50,50"
    'For Each cell In Worksheets("transactionLog").Range("A2:E" & lastRow)
    '    Me.ListBox1.AddItem cell.Value
    'Next cell
    
    ListBox1.RowSource = "transactionLog!A2:E" & lastRow
    'show total number of records
    Me.lblNumRecords.Caption = Me.ListBox1.ListCount
    numRecords = Me.ListBox1.ListCount
    
End Sub
