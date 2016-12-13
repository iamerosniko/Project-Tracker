VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BL_BW_BTExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Sub FillToExcel(ByRef WorkSheet As Integer, ByRef isProjectOrIncident As Boolean, ByRef wb As Workbook)
On Error GoTo err:

    Dim con As New DL_DA_Generic
    Dim row As Integer, col As Integer
    Dim colProj As Collection
    Dim rsDetails As ADODB.Recordset
    Dim i As Integer, ctr As Integer
    Dim temp As String
    Dim colr As Variant
    
    If isProjectOrIncident = True Then
        'Set rsDetails = con.GetData(CMD_EXCELProjDetails & " where BT_ProjDetails.[type]='Project Details'")
        Set rsDetails = con.GetData(CMD_EXCELProjDetails & " and BT_ProjDetails.[type]='Project Details'")
        
        row = 6
    Else
        Set rsDetails = con.GetData(CMD_EXCELProjDetails & " and BT_ProjDetails.[type]='Incident'")
        row = 2
    End If
    col = 1
    
    Do Until rsDetails.EOF
        Set colProj = RecordValueToCollection(con.GetData(CMD_EXCELProjectNames & rsDetails.Fields(12).value), 0)
        ctr = rsDetails.Fields.count - 1
        If Not temp = colProj(1) Then
            wb.Sheets(WorkSheet).Range(wb.Sheets(WorkSheet).Cells(row, 1), wb.Sheets(WorkSheet).Cells(row, ctr)).Merge
            With wb.Sheets(WorkSheet).Cells(row, 1)
                .value = colProj(1)
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = RGB(174, 174, 174)
            End With
            temp = colProj(1)
            row = row + 1
        End If
        If rsDetails.Fields(0).value = "COMPLETE" Then
            colr = RGB(146, 208, 80)
        Else
            colr = RGB(255, 255, 80)
        End If
        With wb.Sheets(WorkSheet).Range(wb.Sheets(WorkSheet).Cells(row, 1), wb.Sheets(WorkSheet).Cells(row, ctr))
            .Font.Size = 10

        End With
        
        
        For i = 0 To rsDetails.Fields.count - 2
            wb.Sheets(WorkSheet).Cells(row, i + 1).value = rsDetails.Fields(i).value
        Next i
        row = row + 1
        rsDetails.MoveNext
    Loop
    Exit Sub
err:
    Exit Sub
End Sub

Public Sub initiateExport()
    Dim app As New Excel.Application
    Dim wb As Excel.Workbook
    Dim file As String
    Dim path As String
    
    file = Format(Now, "mmddyyyyhhmmss")
    path = Sys_FileDialog.saveAs
    If Not Len(path) = 0 Then
        Dim CMD_EXCELTEMPLATE As String
        CMD_EXCELTEMPLATE = CurrentProject.path & "\Template.xlsx"
        Set wb = app.Workbooks.Open(CMD_EXCELTEMPLATE, , True)
        Set wb2 = wb
        FillToExcel 1, True, wb
        FillToExcel 2, False, wb
        
        wb.saveAs filename:=path & "\" & file
        Shell "C:\Windows\explorer.exe /select, """ & path & "\" & file & ".xlsx", vbMaximizedFocus
        app.Quit
    End If
End Sub
