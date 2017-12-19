Sub update()
'
' update 宏
'

    Dim updating_date As String 'sheet1的名字
    Dim netvalue_date As Date '最新净值日期
    Dim last_netvalue_date As Date '上一期净值日期
    Dim nvd_column As String '最新净值日期所在列
    Dim lnvd_column As String '上一期净值日期所在列
    Dim start_row As Integer, end_row As Integer '数据开始和结束行号
    Dim fd_name As String, qt_name As String, CTA_name As String '待更新的三个表名称
    Dim fund_book As Workbook, quant_book As Workbook, CTA_book As Workbook
   
    updating_date = "0920"
    netvalue_date = #9/15/2017#
    last_netvalue_date = #9/8/2017#
    nvd_column = "DR"
    lnvd_column = "DQ"
    start_row = 3
    end_row = 118
    
    fd_name = "平台基金净值_20170920.xlsm"

'    qt_name = "量化策略净值_20170726.xlsx"
'    CTA_name = "期货策略净值_20170726.xlsx"
    
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
    mypath = ThisWorkbook.Path
    
'    Set fund_book = GetObject(mypath & fd_name)
'    Set quant_book = GetObject(mypath & qt_name)
'    Set CTA_book = GetObject(mypath & CTA_name)
'
'    Windows(fund_book.Name).Visible = True
'    Windows(quant_book.Name).Visible = True
'    Windows(CTA_book.Name).Visible = True
'    Application.DisplayAlerts = True
'    Application.ScreenUpdating = True

    Set fund_book = Workbooks(fd_name)
'    Set quant_book = Workbooks(qt_name)
'    Set CTA_book = Workbooks(CTA_name)
        
    Dim fd_namearray() As String, qt_namearray() As String, CTA_namearray() As String
    ReDim fd_namearray(1 To fund_book.Sheets.Count)
'    ReDim qt_namearray(1 To quant_book.Sheets.Count)
'    ReDim CTA_namearray(1 To CTA_book.Sheets.Count)
    
        
    Dim i As Integer
    
    For i = 1 To fund_book.Sheets.Count
        fd_namearray(i) = fund_book.Sheets(i).Name
    Next i
'
'    For i = 1 To quant_book.Sheets.Count
'        qt_namearray(i) = quant_book.Sheets(i).Name
'    Next i
'
'    For i = 1 To CTA_book.Sheets.Count
'        CTA_namearray(i) = CTA_book.Sheets(i).Name
'    Next i
'
    
    'sheets(updating_date).Select
    
    '更新平台基金表
    Dim local_end_row As Integer
    For Each xlsheet In fund_book.Sheets
        xlsheet.Visible = xlSheetVisible
    Next
    fund_book.Sheets("总表").Range("B1") = netvalue_date
    For i = start_row To end_row
        
        If IsError(Application.Match(Cells(i, 2), fd_namearray, 0)) Then
            Debug.Print "no such sheet!" & i
            GoTo continue
        Else
            Debug.Print "updating.." & i
            local_end_row = fund_book.Sheets(CStr(Cells(i, 2))).UsedRange.Rows.Count
            If fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row, 1) = last_netvalue_date Then
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 1) = netvalue_date
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 2) = Range(nvd_column & i)
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row, 3).Copy
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 3).PasteSpecial Paste:=xlPasteFormulas
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row, 4).Copy
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 4).PasteSpecial Paste:=xlPasteFormulas
                fund_book.Sheets(CStr(Cells(i, 2))).Visible = xlSheetHidden
            ElseIf fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row, 1) = netvalue_date Then
                GoTo continue
            Else
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 1) = last_netvalue_date
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 2) = Range(lnvd_column & i)
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row, 3).Copy
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 3).PasteSpecial Paste:=xlPasteFormulas
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row, 4).Copy
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 4).PasteSpecial Paste:=xlPasteFormulas
                
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 2, 1) = netvalue_date
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 2, 2) = Range(lnvd_column & i)
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 3).Copy
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 2, 3).PasteSpecial Paste:=xlPasteFormulas
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 1, 4).Copy
                fund_book.Sheets(CStr(Cells(i, 2))).Cells(local_end_row + 2, 4).PasteSpecial Paste:=xlPasteFormulas
                
            End If
          End If
          
continue:
    Next i
    
    '剩余sheet2中内容
    fund_book.Sheets("sheet2").Range("E1") = netvalue_date
    For i = 2 To 7
        local_end_row = fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).UsedRange.Rows.Count
        If (fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Cells(local_end_row, 1) <> fund_book.Sheets("sheet2").Range("F" & i)) Then
            Debug.Print "updating.."
            fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Cells(local_end_row + 1, 1) = fund_book.Sheets("sheet2").Range("F" & i)
            fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Cells(local_end_row, 2).Copy
            fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Cells(local_end_row + 1, 2).PasteSpecial Paste:=xlPasteFormulas
            fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Cells(local_end_row, 3).Copy
            fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Cells(local_end_row + 1, 3).PasteSpecial Paste:=xlPasteFormulas
            fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Cells(local_end_row, 4).Copy
            fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Cells(local_end_row + 1, 4).PasteSpecial Paste:=xlPasteFormulas
            fund_book.Sheets(CStr(fund_book.Sheets("sheet2").Range("A" & i))).Visible = xlSheetHidden
        Else
            Debug.Print "no need to update.." & i
        End If

    Next i
Debug.Print "Finish!"
End Sub
