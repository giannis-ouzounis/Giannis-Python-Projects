Attribute VB_Name = "Module1"
Sub CleanAndUpdateGlossary()
' DESCRIPTION
    MsgBox "This macro processes an Excel glossary file. It performs the following tasks:" & vbCrLf & _
           "1. Cleans and updates the 'GLOSSARY' sheet by deleting unnecessary columns." & vbCrLf & _
           "2. Generates updated bilingual files based on translated and non-translated acronyms." & vbCrLf & _
           "3. Saves the processed files in an 'Output' folder within the selected directory.", vbInformation, "Macro Description"
    Dim ws As Worksheet
    Dim col As Integer
    Dim lastCol As Integer
    Dim header As String
    Dim folderPath As String
    Dim filePath As String
    Dim newFileName As String
    Dim outputFolder As String
    Dim fileName As String
    Dim wb As Workbook
    Dim fDialog As FileDialog
    Dim file As String
    
    ' Prompt user to select a file
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select an input XLSX File"
        .Filters.Add "Excel Files", "*.xlsx", 1
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "No file selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Extract folder path and file name from the filePath
    folderPath = Left(filePath, InStrRev(filePath, "\"))
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    
    ' Create an "Output" folder within the selected folder
    outputFolder = folderPath & "Output\"
    If Dir(outputFolder, vbDirectory) = "" Then
        MkDir outputFolder
    End If
    
    ' Open the selected workbook
    Set wb = Workbooks.Open(filePath)
        
    ' Process the workbook with the three subroutines
    Call CleanAndUpdateGlossaryAllTerms(wb, Left(fileName, InStrRev(fileName, ".") - 1), outputFolder)
    
    Set wb = Workbooks.Open(filePath)
    Call CleanAndUpdateGlossaryTranslatedAcronyms(wb, Left(fileName, InStrRev(fileName, ".") - 1), outputFolder)
    
    Set wb = Workbooks.Open(filePath)
    Call CleanAndUpdateGlossaryNOTTranslatedAcronyms(wb, Left(fileName, InStrRev(fileName, ".") - 1), outputFolder)
    
    MsgBox "File has been processed and saved in the Output folder."
End Sub

Sub CleanAndUpdateGlossaryAllTerms(wb As Workbook, fileName As String, outputFolder As String)
    ' Code for processing all terms
    Dim ws As Worksheet
    Dim col As Integer
    Dim lastCol As Integer
    Dim header As String
    
    ' Step 1: Delete all sheets except "GLOSSARY"
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name <> "GLOSSARY" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' Unhide any hidden rows and columns in the "GLOSSARY" sheet
    Set ws = wb.Sheets("GLOSSARY")
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    
    ' Step 2: Remove columns with "definition" or "acronym" in the header in the "GLOSSARY" sheet
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For col = lastCol To 1 Step -1
        header = LCase(ws.Cells(1, col).Value)
        If InStr(header, "definition") > 0 Or InStr(header, "acronym") > 0 Then
            ws.Columns(col).Delete
        End If
    Next col
    
    ' Step 3: Ensure the first cell (A1) is always "en_US"
    ws.Cells(1, 1).Value = "en_US"
    
    ' Step 4: Loop through column headers and update with language codes
    For col = 2 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        header = ws.Cells(1, col).Value
        Select Case header
            Case "Italian (IT)": ws.Cells(1, col).Value = "it_IT"
            Case "German (DE)": ws.Cells(1, col).Value = "de_DE"
            Case "French (FR)": ws.Cells(1, col).Value = "fr_FR"
            Case "Canadian French (FR-CA)": ws.Cells(1, col).Value = "fr_CA"
            Case "Spanish (ES)": ws.Cells(1, col).Value = "es_ES"
            Case "Latin American Spanish (ES-LA)": ws.Cells(1, col).Value = "es_LA"
            Case "Portuguese (Portugal) (PT)": ws.Cells(1, col).Value = "pt_PT"
            Case "Brazilian Portuguese (PT-BR)": ws.Cells(1, col).Value = "pt_BR"
            Case "Traditional Chinese (CN)": ws.Cells(1, col).Value = "zh_CN"
            Case "Indonesian Bahasa": ws.Cells(1, col).Value = "id_ID"
            Case "Vietnamese (VN)": ws.Cells(1, col).Value = "vi_VN"
            Case "Greek (GR)": ws.Cells(1, col).Value = "el_GR"
            Case "Bulgarian (BU)": ws.Cells(1, col).Value = "bg_BG"
            Case "Romanian (RO)": ws.Cells(1, col).Value = "ro_RO"
            Case "Korean (KR)": ws.Cells(1, col).Value = "ko_KR"
            Case "Turkish (TR)": ws.Cells(1, col).Value = "tr_TR"
            Case "Slovenian (SI)": ws.Cells(1, col).Value = "sl_SL"
            Case "Hebrew (HE)": ws.Cells(1, col).Value = "he_IL"
            Case "Czech (CZ)": ws.Cells(1, col).Value = "cs_CZ"
            Case "Polish (PL)": ws.Cells(1, col).Value = "pl_PL"
            Case "Ukrainian (UA)", "Ukranian (UA)": ws.Cells(1, col).Value = "uk_UA"
        End Select
    Next col
    
    ' Step 5: Save the updated workbook with "_updated" at the end of the file name
    newFileName = outputFolder & fileName & "_updated.xlsx"
    
    ' Save the workbook as a new XLSX file without macros
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=newFileName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
End Sub

Sub CleanAndUpdateGlossaryTranslatedAcronyms(wb As Workbook, fileName As String, outputFolder As String)
    ' Your existing code for processing translated acronyms
    Dim ws As Worksheet
    Dim col As Variant
    Dim lastCol As Integer
    Dim header As String
    Dim translatedFolder As String
    Dim newWb As Workbook
    Dim acronymCols As Collection
    
    ' Create a "Translated acronyms" subfolder within the Output folder
    translatedFolder = outputFolder & "Translated acronyms\"
    If Dir(translatedFolder, vbDirectory) = "" Then
        MkDir translatedFolder
    End If
    
    ' Step 1: Delete all sheets except "GLOSSARY"
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name <> "GLOSSARY" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' Unhide any hidden rows and columns in the "GLOSSARY" sheet
    Set ws = wb.Sheets("GLOSSARY")
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    
    ' Step 2: Identify and collect columns with "acronym" in the header
    Set acronymCols = New Collection
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        header = Trim(LCase(ws.Cells(1, col).Value))
        If InStr(header, "acronym") > 0 Then
            acronymCols.Add col
        End If
    Next col
    
    ' Step 3: Rename columns according to the provided mappings
    For Each col In acronymCols
        header = Trim(LCase(ws.Cells(1, col).Value))
        If ws.Cells(1, col).Value = "Turkish acronym" Then
            ws.Cells(1, col).Value = "tr_TR"
        ElseIf ws.Cells(1, col).Value = "ID Acronym" Then
            ws.Cells(1, col).Value = "id_ID"
        End If
        
        Select Case header
            Case "en acronym": ws.Cells(1, col).Value = "en_US"
            Case "it acronym": ws.Cells(1, col).Value = "it_IT"
            Case "de acronym": ws.Cells(1, col).Value = "de_DE"
            Case "fr acronym": ws.Cells(1, col).Value = "fr_FR"
            Case "fr-ca acronym": ws.Cells(1, col).Value = "fr_CA"
            Case "es acronym": ws.Cells(1, col).Value = "es_ES"
            Case "es-la acronym": ws.Cells(1, col).Value = "es_LA"
            Case "pt acronym": ws.Cells(1, col).Value = "pt_PT"
            Case "pt-br acronym": ws.Cells(1, col).Value = "pt_BR"
            Case "cn acronym": ws.Cells(1, col).Value = "zh_CN"
            Case "id acronym": ws.Cells(1, col).Value = "id_ID"
            Case "vn acronym": ws.Cells(1, col).Value = "vi_VN"
            Case "gr acronym": ws.Cells(1, col).Value = "el_GR"
            Case "bu acronym": ws.Cells(1, col).Value = "bg_BG"
            Case "ro acronym": ws.Cells(1, col).Value = "ro_RO"
            Case "kr acronym": ws.Cells(1, col).Value = "ko_KR"
            Case "tr acronym": ws.Cells(1, col).Value = "tr_TR"
            Case "si acronym": ws.Cells(1, col).Value = "sl_SL"
            Case "he acronym": ws.Cells(1, col).Value = "he_IL"
            Case "cz acronym": ws.Cells(1, col).Value = "cs_CZ"
            Case "pl acronym": ws.Cells(1, col).Value = "pl_PL"
            Case "ua acronym": ws.Cells(1, col).Value = "uk_UA"
        End Select
    Next col
    
    ' Step 4: Create new bilingual XLSX files
    For Each col In acronymCols
        If ws.Cells(1, col).Value <> "en_US" Then
            Set newWb = Workbooks.Add
            With newWb.Sheets(1)
                .Cells(1, 1).Value = "en_US"
                .Cells(1, 2).Value = ws.Cells(1, col).Value
                
                ' Copy data from en_US and the selected language column
                .Range(.Cells(2, 1), .Cells(ws.Rows.Count, 1)).Value = ws.Range(ws.Cells(2, acronymCols(1)), ws.Cells(ws.Rows.Count, acronymCols(1))).Value
                .Range(.Cells(2, 2), .Cells(ws.Rows.Count, 2)).Value = ws.Range(ws.Cells(2, col), ws.Cells(ws.Rows.Count, col)).Value
            End With
            
            ' Save the new file in the Translated acronyms subfolder
            If ws.Cells(1, col).Value = "tr_TR" Then
                newFileName = translatedFolder & fileName & "_tr_TR.xlsx"
            ElseIf ws.Cells(1, col).Value = "id_ID" Then
                newFileName = translatedFolder & fileName & "_id_ID.xlsx"
            Else
                newFileName = translatedFolder & fileName & "_" & ws.Cells(1, col).Value & ".xlsx"
            End If
            
            Application.DisplayAlerts = False
            newWb.SaveAs fileName:=newFileName, FileFormat:=xlOpenXMLWorkbook
            newWb.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
    Next col
    
    ' Close the original workbook without saving changes
    wb.Close SaveChanges:=False
End Sub

Sub CleanAndUpdateGlossaryNOTTranslatedAcronyms(wb As Workbook, fileName As String, outputFolder As String)
    ' Code for processing NOT translated acronyms
    Dim ws As Worksheet
    Dim col As Variant
    Dim lastCol As Integer
    Dim header As String
    Dim notTranslatedFolder As String
    Dim newWb As Workbook
    Dim acronymCols As Collection
    
    ' Create a "NOT translated acronyms" subfolder within the Output folder
    notTranslatedFolder = outputFolder & "NOT translated acronyms\"
    If Dir(notTranslatedFolder, vbDirectory) = "" Then
        MkDir notTranslatedFolder
    End If
    
    ' Step 1: Delete all sheets except "GLOSSARY"
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name <> "GLOSSARY" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' Unhide any hidden rows and columns in the "GLOSSARY" sheet
    Set ws = wb.Sheets("GLOSSARY")
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    
    ' Step 2: Identify and collect columns with "acronym" in the header
    Set acronymCols = New Collection
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        header = Trim(LCase(ws.Cells(1, col).Value))
        If InStr(header, "acronym") > 0 Then
            acronymCols.Add col
        End If
    Next col
    
    ' Step 3: Rename columns according to the provided mappings
    For Each col In acronymCols
        header = Trim(LCase(ws.Cells(1, col).Value))
        If ws.Cells(1, col).Value = "Turkish acronym" Then
            ws.Cells(1, col).Value = "tr_TR"
        ElseIf ws.Cells(1, col).Value = "ID Acronym" Then
            ws.Cells(1, col).Value = "id_ID"
        End If
        
        Select Case header
            Case "en acronym": ws.Cells(1, col).Value = "en_US"
            Case "it acronym": ws.Cells(1, col).Value = "it_IT"
            Case "de acronym": ws.Cells(1, col).Value = "de_DE"
            Case "fr acronym": ws.Cells(1, col).Value = "fr_FR"
            Case "fr-ca acronym": ws.Cells(1, col).Value = "fr_CA"
            Case "es acronym": ws.Cells(1, col).Value = "es_ES"
            Case "es-la acronym": ws.Cells(1, col).Value = "es_LA"
            Case "pt acronym": ws.Cells(1, col).Value = "pt_PT"
            Case "pt-br acronym": ws.Cells(1, col).Value = "pt_BR"
            Case "cn acronym": ws.Cells(1, col).Value = "zh_CN"
            Case "id acronym": ws.Cells(1, col).Value = "id_ID"
            Case "vn acronym": ws.Cells(1, col).Value = "vi_VN"
            Case "gr acronym": ws.Cells(1, col).Value = "el_GR"
            Case "bu acronym": ws.Cells(1, col).Value = "bg_BG"
            Case "ro acronym": ws.Cells(1, col).Value = "ro_RO"
            Case "kr acronym": ws.Cells(1, col).Value = "ko_KR"
            Case "tr acronym": ws.Cells(1, col).Value = "tr_TR"
            Case "si acronym": ws.Cells(1, col).Value = "sl_SL"
            Case "he acronym": ws.Cells(1, col).Value = "he_IL"
            Case "cz acronym": ws.Cells(1, col).Value = "cs_CZ"
            Case "pl acronym": ws.Cells(1, col).Value = "pl_PL"
            Case "ua acronym": ws.Cells(1, col).Value = "uk_UA"
        End Select
    Next col
    
    ' Step 4: Create new bilingual XLSX files
    For Each col In acronymCols
        If ws.Cells(1, col).Value <> "en_US" Then
            Set newWb = Workbooks.Add
            With newWb.Sheets(1)
                .Cells(1, 1).Value = "en_US"
                .Cells(1, 2).Value = ws.Cells(1, col).Value
                
                ' Copy data from en_US and the selected language column
                .Range(.Cells(2, 1), .Cells(ws.Rows.Count, 1)).Value = ws.Range(ws.Cells(2, acronymCols(1)), ws.Cells(ws.Rows.Count, acronymCols(1))).Value
                .Range(.Cells(2, 2), .Cells(ws.Rows.Count, 2)).Value = ws.Range(ws.Cells(2, col), ws.Cells(ws.Rows.Count, col)).Value
                
                ' Copy contents of column A to column B excluding row 1
                .Range(.Cells(2, 2), .Cells(ws.Rows.Count, 2)).Value = .Range(.Cells(2, 1), .Cells(ws.Rows.Count, 1)).Value
            End With
            
            ' Save the new file in the NOT translated acronyms subfolder
            newFileName = notTranslatedFolder & fileName & "_" & ws.Cells(1, col).Value & ".xlsx"
            
            Application.DisplayAlerts = False
            newWb.SaveAs fileName:=newFileName, FileFormat:=xlOpenXMLWorkbook
            newWb.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
    Next col
    
    ' Close the original workbook without saving changes
    wb.Close SaveChanges:=False
End Sub

