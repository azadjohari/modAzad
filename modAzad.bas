Attribute VB_Name = "modAzad"
Sub Filter_Records(ByVal sheetname As String, ByVal rng As String, ByVal fieldNo As Long, ByVal criteria As String)
    Sheets(sheetname).Select
    Range(rng).AutoFilter Field:=fieldNo, Criteria1:=criteria
    'to filter blanks, set criteria to <>
End Sub

Sub Filter_Records_By_Color(ByVal sheetname As String, ByVal rng As String, ByVal fieldNo As Long, Optional ByVal rgb_code As Long = 16777215, Optional ByVal NoFill As Boolean = False)
    Sheets(sheetname).Select
    
    If NoFill Then
        Range(rng).AutoFilter Field:=fieldNo, Operator:=xlFilterNoFill
    Else
        Range(rng).AutoFilter Field:=fieldNo, Criteria1:=rgb_code, Operator:=xlFilterCellColor
    End If
    
End Sub
Sub Filter_Clear(ByVal sheetname As String)
    Sheets(sheetname).Select
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    If ActiveSheet.AutoFilterMode Then
        Cells.AutoFilter
    End If
End Sub

Sub Filter_Copy_PasteSpecial(ByVal sheetNameCopy As String, ByVal rng As String, ByVal sheetNamePaste As String)
    
    
        Sheets(sheetNamePaste).Cells.Clear
    With Sheets(sheetNameCopy)
            
        Dim r As Range
        Set r = .Range(rng).SpecialCells(xlCellTypeVisible)
        r.Copy
            
        Sheets(sheetNamePaste).Range("A1").Select
        Sheets(sheetNamePaste).Paste
        Application.CutCopyMode = False
    End With
End Sub

Sub Filter_Copy_Filtered_Range_And_Paste_To_New_Sheet(ByRef ws As Worksheet)
    ActiveSheet.AutoFilter.Range.Copy
    ws.Select
    Range("A1").Select
    ws.Paste
End Sub
Sub Filter_And_Paste_Result_To_New_Sheet( _
        ByVal new_sheet_name As String, ByRef sheet_to_filter As Worksheet, _
        ByVal range_included_in_filter As Range, ByVal column_to_filter As Double, _
        ByVal filter_criteria As String)
    Sheets_Clear_Or_Create_New (new_sheet_name)
    
    Filter_Clear (sheet_to_filter.Name)
    Filter_Records sheet_to_filter.Name, range_included_in_filter.Address(False, False), column_to_filter, filter_criteria
    Filter_Copy_Filtered_Range_And_Paste_To_New_Sheet Sheets(new_sheet_name)
End Sub
'Sub Range_Copy_Paste(ByVal sheetNameCopy As String, ByVal rng As String, ByVal sheetNamePaste As String, ByVal pasteLocation As String)
'
'    With Sheets(sheetNameCopy)
'
'        Dim r As Range
'        Set r = .Range(rng)
'        r.Copy
'
'        .Range(pasteLocation).Select
'        Sheets(sheetNamePaste).Paste
'        Application.CutCopyMode = False
'    End With
'End Sub
Sub Range_Copy_Paste(ByRef copy_from As Range, ByRef sheet_to_paste As Worksheet, _
                    ByRef paste_location_r1c1 As String)
    copy_from.Copy
    sheet_to_paste.Select
    Range(paste_location_r1c1).Select
    sheet_to_paste.Paste
    Application.CutCopyMode = False
End Sub


Sub Range_Clear(ByVal sheetname As String, ByVal rng As String, Optional ByVal clearAll As Boolean = False, Optional ByVal clearFormatAndContents As Boolean = False)
    With Sheets(sheetname)
        If clearAll Then
            .Range(rng).Clear
        ElseIf clearFormatAndContents Then
            .Range(rng).ClearFormats
            .Range(rng).ClearContents
        Else
            .Range(rng).ClearContents
        End If
    End With
End Sub

Function Range_Count(ByVal sheetname As String, ByVal data_start_col As String, ByVal data_start_row As Long) As Long
    Dim ret As Long
    Sheets(sheetname).Select
    With Sheets(sheetname)
    
        If .Range(data_start_col & data_start_row) <> "" Then
            .Range(data_start_col & data_start_row).Select
            If .Range(data_start_col & (data_start_row + 1)) <> "" Then
                .Range(Selection, Selection.End(xlDown)).Select
            End If
            ret = Selection.count
        Else
            ret = 0
        End If
    End With
    
    Range_Count = ret
End Function

Function Range_Selection_Down(ByVal sheetname As String, ByVal data_start_col As String, ByVal data_start_row As Long) As Boolean
    Dim ret As Boolean
    Sheets(sheetname).Select
    With Sheets(sheetname)
    
        If .Range(data_start_col & data_start_row) <> "" Then
            .Range(data_start_col & data_start_row).Select
            If .Range(data_start_col & (data_start_row + 1)) <> "" Then
                .Range(Selection, Selection.End(xlDown)).Select
            End If
            ret = True
        Else
            ret = False
        End If
    End With
    
    Range_Selection_Down = ret
    
End Function

Sub Range_Update_Value(ByRef wk As Workbook, ByRef sht As Worksheet, _
                    ByRef affected_range As Range, ByRef value_to_update As String, _
                    Optional ByRef screenupdate As Boolean = False)
    modAzad.Common_Activate_WK_And_SHT wk, sht
    Application.ScreenUpdating = screenupdate
    DoEvents
    affected_range = value_to_update
    Application.ScreenUpdating = False
End Sub


Function Range_Get_Value(ByVal search_item As String, ByVal search_ws As Worksheet, ByVal search_range As String, ByVal value_column As String) As String

    Dim ret As String
    ret = ""
    Dim searchResult As Variant
    
    Set searchResult = search_ws.Range(search_range).Find(What:="" & search_item & "")
    
    If Not searchResult Is Nothing Then
        ret = search_ws.Range(value_column & searchResult.Row)
    Else
        ret = "Not found"
    End If
    Range_Get_Value = ret
End Function

Sub Range_Draw_Line(ByVal sheetname As String, ByVal rng As String)
    Sheets(sheetname).Select
    Range(rng).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub
Sub Range_Remove_Empty_Row(ByVal sheetname As String, ByVal referenceRange As String)
    Sheets(sheetname).Select
    Range(referenceRange).Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
End Sub

Sub Range_Remove_Duplicates(ByVal sheetname As String, ByVal rangeToRemoveDuplicates As Variant)
    Sheets(sheetname).Select
    If TypeName(rangeToRemoveDuplicates) = "Range" Then
        rangeToRemoveDuplicates.RemoveDuplicates Columns:=1, Header:=xlNo
    Else
        Range(rangeToRemoveDuplicates).RemoveDuplicates Columns:=1, Header:=xlNo
    End If
'    Range("C5").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Sheets.Add After:=Sheets(Sheets.Count)
'    Range("A1").Select
'    ActiveSheet.Paste
'    Application.CutCopyMode = False
'    ActiveSheet.Range("$A$1:$A$1450").RemoveDuplicates Columns:=1, Header:=xlNo
End Sub

Sub Range_Sort(ByRef sheet As Worksheet, ByRef rng As Range, _
            Optional ByVal sort_order As XlSortOrder = xlAscending, _
            Optional ByVal data_option As XlSortDataOption = xlSortNormal, _
            Optional ByVal sort_on As XlSortOn = xlSortOnValues)
            
    sheet.Sort.SortFields.Clear
    sheet.Sort.SortFields.Add Key:=rng, SortOn:=sort_on, order:=sort_order, DataOption:= _
        data_option
    With sheet.Sort
        .SetRange rng
        .Header = xlNo
        .MatchCase = False
        .orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Range_Create_List(ByRef wk As Workbook, ByRef sht As Worksheet, ByRef affected_range As Range, ByRef list_value As String)
    modAzad.Common_Activate_WK_And_SHT wk, sht
    affected_range.Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=list_value
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Function Workbook_Is_Active(ByRef wk As Workbook) As Boolean
    If wk.Name = ActiveWorkbook.Name Then
        Workbook_Is_Active = True
    Else
        Workbook_Is_Active = False
    End If
End Function


Sub App_Init(ByVal status As Boolean)
    Application.ScreenUpdating = status
    Application.DisplayAlerts = status
    Application.AskToUpdateLinks = status
End Sub

Sub App_StatusBar(ByVal msg As String)
    Application.StatusBar = msg
End Sub

Function Sheets_Is_Available(ByVal sSheetName As String) As Boolean

    On Error Resume Next
    Dim oSheet As Excel.Worksheet

    Set oSheet = ActiveWorkbook.Sheets(sSheetName)
    Sheets_Is_Available = IIf(oSheet Is Nothing, False, True)

End Function

Sub Sheets_Copy_New(ByVal sheetname_to_copy As String, ByVal new_copied_sheet_name As String, ByVal new_sheet_position_after As Double)
    Sheets(sheetname_to_copy).Select
    Sheets(sheetname_to_copy).Copy After:=Sheets(new_sheet_position_after)
    Sheets(new_sheet_position_after + 1).Select
    Sheets(new_sheet_position_after + 1).Name = new_copied_sheet_name
    
End Sub

Sub Sheets_Clear_Or_Create_New(ByVal sheetname As String)
    If Sheets_Is_Available(sheetname) Then
        Sheets(sheetname).Cells.Clear
    Else
        Sheets.Add After:=Sheets(Sheets.count)
        ActiveSheet.Name = sheetname
    End If
End Sub

Sub Sheets_Delete_And_Create_New(ByVal sheetname As String, Optional ByVal color_tab As Boolean = False, _
                    Optional ByVal color_code As Double = 65535)
    If Sheets_Is_Available(sheetname) Then
        Sheets_Delete (sheetname)
    End If
    Sheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.Name = sheetname
    
    If color_tab Then
        With Sheets(sheetname).Tab
            .Color = color_code
            .TintAndShade = 0
        End With
    End If
End Sub

Sub Sheets_Visibility(ByRef ws As Worksheet, ByVal status As Boolean)
    ws.visible = status
End Sub

Sub Sheets_Delete(ByVal sheetname As String)
    If modAzad.Sheets_Is_Available(sheetname) Then
        Sheets(sheetname).Delete
    End If
End Sub

Sub Sheets_Zoom(ByRef sheet As Worksheet, ByVal percentage As Integer)
    sheet.Activate
    ActiveWindow.Zoom = percentage
End Sub
Function Sheets_Is_Active(ByRef sht As Worksheet) As Boolean
    If sht.Name = ActiveSheet.Name Then
        Sheets_Is_Active = True
    Else
        Sheets_Is_Active = False
    End If
End Function


Function Shapes_Is_Available(ByRef ws As Worksheet, ByVal shape_name As String)
    ws.Activate
    On Error Resume Next
    Dim oShape As Object

    Set oSheet = ActiveSheet.Shapes(shape_name)
    Shapes_Is_Available = IIf(oSheet Is Nothing, False, True)
End Function
Sub Shapes_Delete_And_Create_New_Textbox(ByRef ws As Worksheet, ByVal shape_name As String, _
    Optional ByVal txt As String = "", Optional ByVal orientation As MsoTextOrientation = msoTextOrientationHorizontal, _
    Optional ByVal pos_top = 0, Optional ByVal pos_left = 0, Optional ByVal size_height = 100, _
    Optional ByVal size_width = 100, Optional ByVal fill_visiblity As MsoTriState = msoFalse, _
    Optional ByVal line_visiblity As MsoTriState = msoFalse)

    If Shapes_Is_Available(ws, shape_name) Then
        ActiveSheet.Shapes(shape_name).Delete
    End If
    
    ws.Activate
    ActiveSheet.Shapes.AddTextbox(orientation, pos_left, pos_top, size_width, size_height).Select
    Selection.ShapeRange.Name = shape_name
    Selection.ShapeRange.Fill.visible = fill_visiblity
    Selection.ShapeRange.Line.visible = line_visiblity
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = txt
End Sub
Sub Shapes_Position(ByRef ws As Worksheet, ByVal shape_name As String, ByVal pos_top As Long, ByVal pos_left As Long)
    ws.Activate
    With ActiveSheet.Shapes(shape_name)
        .Left = pos_left
        .top = pos_top
    End With
End Sub
Sub Shapes_Popup_Info(ByVal shape_name As String)
    MsgBox "Info for shape_name : " & shape_name & vbCr _
    & "Top = " & ActiveSheet.Shapes(shape_name).top & vbCr _
    & "Left = " & ActiveSheet.Shapes(shape_name).Left & vbCr _
    & "Height = " & ActiveSheet.Shapes(shape_name).Height & vbCr _
    & "Width = " & ActiveSheet.Shapes(shape_name).Width
End Sub
Sub Common_Activate_WK_And_SHT(ByRef wk As Workbook, ByRef sht As Worksheet)
    If modAzad.Workbook_Is_Active(wk) = False Then
        wk.Activate
    End If
    If modAzad.Sheets_Is_Active(sht) = False Then
        sht.Activate
    End If
End Sub
