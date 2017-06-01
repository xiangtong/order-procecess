'-----------各个表格的名称和格式都不能改变----------------------
'-------定义全局变量------------

Const ProductSheet = "Product"
Const OrderSheet = "Order"
Const PickSheet = "Pick"
Const PXSheet = "4px"
Dim ResultSheet, USPackSheet, USPickSheet, SZPacksheet, SZPickSheet, SZ4PXSheet As String

Dim Ordercount, Productcount, USOrdercount, SZOrdercount As Integer

Dim Sale_number, Product_SKU As String
Public Ebay_account As String

Sub Main()

    '----------------Main sub-------------    '----------------active current workbook-------
    Workbookname = ThisWorkbook.Name
    Workbooks(Workbookname).Activate


    '--------------判断是否存在基础表格和是否选择了Ebay账号,如果不存在或未选择，则提示错误和退出程序-----------
    Call prerequisite

    '--------------set all result sheets name based on current date-----------
    Call set_result_sheets_name

    '------------copy order sheet to a result sheet-------------------
    Call copy_sheet(OrderSheet, ResultSheet)

    '------------insert 4 columns into result sheet-------------------
    Call insert_column

    '-------------获得各种则数量，以便下面循环遍历时使用---------
    Ordercount = Getrowcount(OrderSheet)
    Productcount = Getrowcount(ProductSheet)
    'MsgBox Ordercount & vbCrLf & Productcount

    '--------pre process: add lacked shipping service------
    Call pre_process_resultsheet(ResultSheet)

    '------Get purchaser and Chinese name from product sheet--------------------------
    Call add_product_info

    '-------insert 1 column to contain merged address--------
    Call insert_col_for_address

    '---------merge addresses------------
    Call mergeaddress

    '--------sorting-----
    Call result_sheet_sorting(ResultSheet, Ordercount, 1)

    '-----------focus to result sheeet-----------
    sheets(ResultSheet).Activate

    '----------add ebay account info to order number
    Call addebayaccount(ResultSheet)

    '------format result sheet
    Call formatresultsheet

    '------split result sheet into US_Pack and SZ_Pack
    Call copy_sheet(ResultSheet, USPackSheet)
    Call copy_sheet(ResultSheet, SZPacksheet)

    '------process US Pack Sheet------
    Call p_us_pack_sheet(USPackSheet)

    '------Create US picking sheet-----
    Call copy_sheet(PickSheet, USPickSheet)
    Call gen_pick_sheet(USPackSheet, USPickSheet, "US")

    '------process SZ Pack Sheet------
    Call p_sz_pack_sheet(SZPacksheet)

    '------Create SZ picking sheet-----
    Call copy_sheet(PickSheet, SZPickSheet)
    Call gen_pick_sheet(SZPacksheet, SZPickSheet, "SZ")

    '------Create SZ shipping table for 4px-----
    Call copy_sheet(PXSheet, SZ4PXSheet)
    Call gen_4px_sheet(SZPacksheet, SZ4PXSheet)

    '----merge cells for us and sz pack tables
    '--------merge same order id--------
    'MsgBox USOrdercount & vbCrLf & SZOrdercount
    Call merge_same_orderid(USPackSheet, USOrdercount, "US")
    Call merge_same_orderid(SZPacksheet, SZOrdercount, "SZ")
    Call formatpacksheet(USPackSheet)
    Call formatpacksheet(SZPacksheet)

    'active the last sheet
    sheets(SZ4PXSheet).Activate
    Range("A1").Select

End Sub


Sub gen_4px_sheet(ByVal sourcesheet As String, ByVal dessheet As String)
    '--------generate table for 4px upload based on SZ pack table-----
    Dim i, j, pickcount As Integer
    Dim price As Long
    Dim ifexist As Boolean
    pickcount = Getrowcount(dessheet) - 1

    For i = 2 To 59
        ifexist = False
        For j = 2 To pickcount
            If sheets(sourcesheet).Cells(i, 1).Value = sheets(dessheet).Cells(j, 1).Value Then
                If sheets(dessheet).Cells(j, 30).Value = "" Then
                    sheets(dessheet).Cells(j, 30).Value = sheets(sourcesheet).Cells(i, 18).Value
                Else
                    sheets(dessheet).Cells(j, 30).Value = sheets(dessheet).Cells(j, 30).Value + sheets(sourcesheet).Cells(i, 18).Value
                End If
                ifexist = True
                Exit For
            End If
        Next j
        If ifexist = False And sheets(sourcesheet).Cells(i, 42).Value = "新加坡小包" Then
            sheets(dessheet).Cells(pickcount, 1).Value = sheets(sourcesheet).Cells(i, 1).Value
            sheets(dessheet).Cells(pickcount, 3).Value = "B1"
            sheets(dessheet).Cells(pickcount, 4).Value = sheets(sourcesheet).Cells(i, 12).Value
            sheets(dessheet).Cells(pickcount, 12).Value = sheets(sourcesheet).Cells(i, 3).Value
            sheets(dessheet).Cells(pickcount, 13).Value = sheets(sourcesheet).Cells(i, 10).Value
            sheets(dessheet).Cells(pickcount, 14).Value = sheets(sourcesheet).Cells(i, 9).Value
            sheets(dessheet).Cells(pickcount, 15).Value = sheets(sourcesheet).Cells(i, 8).Value
            sheets(dessheet).Cells(pickcount, 16).Value = sheets(sourcesheet).Cells(i, 4).Value
            sheets(dessheet).Cells(pickcount, 18).Value = sheets(sourcesheet).Cells(i, 11).Value
            sheets(dessheet).Cells(pickcount, 27).Value = sheets(sourcesheet).Cells(i, 14).Value
            sheets(dessheet).Cells(pickcount, 30).Value = sheets(sourcesheet).Cells(i, 18).Value
            pickcount = pickcount + 1
        End If
    Next i
    '--------下面是申报价格，这个还是要手工填写比较合适。所以注释下面的代码
'    For j = 2 To pickcount
'        MsgBox Sheets(dessheet).Cells(j, 30).Value
'        If Sheets(dessheet).Cells(j, 30).Value <> "" Then
'            price = 50 / Sheets(dessheet).Cells(j, 30).Value
'            Sheets(dessheet).Cells(j, 29).Value = Application.WorksheetFunction.RoundDown(price, 0)
'        End If
'    Next j

    '-----format------
    sheets(dessheet).Activate
    Cells.Select
    With Selection.font
        .Name = "宋体"
        .Size = 10
    End With
    With Selection
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        Rows.AutoFit
    End With

    '------remove "_" from order id----
    'Dim temp As String
    For j = 2 To pickcount
        'temp = sheets(dessheet).Cells(j, 1).Value
        'temp = Right(temp, (Len(temp) - InStrRev(temp, "_")))
        'sheets(dessheet).Cells(j, 1).Value = temp
        sheets(dessheet).Cells(j, 1).Value = LCase(Replace(sheets(dessheet).Cells(j, 1).Value, "_", ""))
    Next j
End Sub

Sub gen_pick_sheet(ByVal sourcesheet As String, ByVal dessheet As String, ByVal sheettype As String)

    '------copy row and caculate summary-----
    Dim i, j, pickcount As Integer
    Dim ifexist As Boolean
    pickcount = Getrowcount(dessheet)

    If sheettype = "US" Then
        For i = 2 To USOrdercount
            ifexist = False
            For j = 3 To pickcount
                If sheets(sourcesheet).Cells(i, 15).Value = sheets(dessheet).Cells(j, 1).Value Then
                    sheets(dessheet).Cells(j, 6).Value = sheets(dessheet).Cells(j, 6).Value + sheets(sourcesheet).Cells(i, 16).Value
                    ifexist = True
                    Exit For
                End If
            Next j
            If ifexist = False Then
                sheets(dessheet).Cells(pickcount, 1).Value = sheets(sourcesheet).Cells(i, 15).Value
                sheets(dessheet).Cells(pickcount, 2).Value = sheets(sourcesheet).Cells(i, 18).Value
                sheets(dessheet).Cells(pickcount, 3).Value = sheets(sourcesheet).Cells(i, 14).Value
                sheets(dessheet).Cells(pickcount, 5).Value = sheets(sourcesheet).Cells(i, 19).Value
                sheets(dessheet).Cells(pickcount, 6).Value = sheets(sourcesheet).Cells(i, 16).Value
                pickcount = pickcount + 1
            End If
        Next i
    ElseIf sheettype = "SZ" Then
        For i = 2 To SZOrdercount
            ifexist = False
            For j = 3 To pickcount
                If sheets(sourcesheet).Cells(i, 17).Value = sheets(dessheet).Cells(j, 1).Value Then
                    sheets(dessheet).Cells(j, 6).Value = sheets(dessheet).Cells(j, 6).Value + sheets(sourcesheet).Cells(i, 18).Value
                    ifexist = True
                    Exit For
                End If
            Next j
            If ifexist = False Then
                sheets(dessheet).Cells(pickcount, 1).Value = sheets(sourcesheet).Cells(i, 17).Value
                sheets(dessheet).Cells(pickcount, 2).Value = sheets(sourcesheet).Cells(i, 15).Value
                sheets(dessheet).Cells(pickcount, 3).Value = sheets(sourcesheet).Cells(i, 14).Value
                sheets(dessheet).Cells(pickcount, 5).Value = sheets(sourcesheet).Cells(i, 16).Value
                sheets(dessheet).Cells(pickcount, 6).Value = sheets(sourcesheet).Cells(i, 18).Value
                pickcount = pickcount + 1
            End If
        Next i
    End If

        '------add mapping relation into the first row
    sheets(dessheet).Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "此分拣单对应的装箱单是： " & sourcesheet
    Call merge_cells("A1:F1")

End Sub


Sub p_us_pack_sheet(ByVal sheetname As String)

    sheets(sheetname).Activate
    'hide useless columns
    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("D:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("Q:R").Select
    Selection.EntireColumn.Hidden = True

    '----change title -----
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "图片"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "账号"

    '-------change column order------
    Columns("C:C").Select
    Selection.cut
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight
    Columns("U:U").Select
    Selection.cut
    Columns("AO:AO").Select
    Selection.Insert Shift:=xlToRight

    '-------insert 2 new columns-----
    Columns("AO:AO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AO1").Select
    ActiveCell.FormulaR1C1 = "实际费用"
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AP1").Select
    ActiveCell.FormulaR1C1 = "发货日期"

    '---delete all rows which do not belong to US--
    Dim k As Integer
    Dim rowrange As String
    For k = Ordercount To 2 Step -1
        If InStr(1, sheets(sheetname).Cells(k, 35).Value, "USPS", vbTextCompare) = 0 Then
            rowrange = CStr(k) & ":" & CStr(k)
            Rows(rowrange).Select
            Selection.Delete Shift:=xlUp
        End If
    Next k

    Columns("AI:AI").ColumnWidth = 15
    USOrdercount = Getrowcount(USPackSheet)


End Sub
Sub p_sz_pack_sheet(ByVal sheetname As String)

    sheets(sheetname).Activate

    '----change title -----
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "订单号"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "买家ID"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "收货人"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "买家电话号码"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "地址"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "城市"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "州"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "邮编"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "国家"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "EBay产品编号"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "产品英文名称"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "SKU"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "数量"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "单价"
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "运单号"

    '-------hide useless columns
    Columns("Z:Z").Select
    Selection.EntireColumn.Hidden = True
    Columns("AJ:AJ").Select
    Selection.EntireColumn.Hidden = True

    '-------change column order------
    Columns("C:C").Select
    Selection.cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("R:R").Select
    Selection.cut
    Columns("O:O").Select
    Selection.Insert Shift:=xlToRight
    Columns("S:S").Select
    Selection.cut
    Columns("P:P").Select
    Selection.Insert Shift:=xlToRight

    '-------insert 2 new columns-----
    Columns("AO:AO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AO1").Select
    ActiveCell.FormulaR1C1 = "实际重量"
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AP1").Select
    ActiveCell.FormulaR1C1 = "运输方式"

    '------sorting for delete
    Call result_sheet_sorting(sheetname, Ordercount, 0)

    '---delete all rows which do not belong to US--
    Dim k As Integer
    Dim rowrange As String
    For k = Ordercount To 2 Step -1
        If InStr(1, sheets(sheetname).Cells(k, 36).Value, "USPS", vbTextCompare) <> 0 Then
            rowrange = CStr(k) & ":" & CStr(k)
            Rows(rowrange).Select
            Selection.Delete Shift:=xlUp
        End If
    Next k

    '---- update current row count after deleting
    SZOrdercount = Getrowcount(sheetname)

    '-----sorting by country and order id
    Call sort_by_country(sheetname)

    '----generate shipping service-----
    Call generate_ship_service(sheetname, SZOrdercount)

    Range("A1").Select
End Sub
Sub generate_ship_service(ByVal sheetname As String, rowcount As Integer)
    sheets(sheetname).Activate
    For k = 2 To rowcount
        If sheets(sheetname).Cells(k, 12).Value = "United States" Then
            sheets(sheetname).Cells(k, 42).Value = "e邮宝"
        ElseIf sheets(sheetname).Cells(k, 12).Value <> "United States" Then
            sheets(sheetname).Cells(k, 42).Value = "新加坡小包"
        End If
    Next k
End Sub
Sub sort_by_country(ByVal sheetname As String)
    Dim keyrange, keyrange2, sortrange As String
    sheets(sheetname).Activate
    keyrange = "L2:L" & CStr(SZOrdercount)
    keyrange2 = "A2:A" & CStr(SZOrdercount)
    sortrange = "A1:AP" & CStr(SZOrdercount)
    ActiveWorkbook.Worksheets(sheetname).sort.SortFields.Clear
    ActiveWorkbook.Worksheets(sheetname).sort.SortFields.Add Key:=Range( _
        keyrange), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(sheetname).sort.SortFields.Add Key:=Range( _
        keyrange2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(sheetname).sort
        .SetRange Range(sortrange)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub prerequisite()
    If SheetExists(PXSheet) = False Or SheetExists(PickSheet) = False Or SheetExists(ProductSheet) = False Or SheetExists(OrderSheet) = Fasle Then
        MsgBox " 缺少""Product表""或者""Order表""或者""Pick表""或者""4px表""，错误，退出计算！"
        End
    End If
    '-------------need select ebay account before caculate---------
    If Ebay_account = "" Then
        MsgBox "请选择一个ebay账号，再点击Go！"
        End
    End If
    'MsgBox Ebay_account
End Sub
Sub set_result_sheets_name()
    '-----set names for all result sheets based on current date---
        Dim m, d As String
        m = Month(Date)
        d = Day(Date)
        If Len(m) = 1 Then
            m = "0" & m
        End If
        If Len(d) = 1 Then
            d = "0" & d
        End If
        ResultSheet = "Total_" & Year(Date) & "-" & m & "-" & d
        USPackSheet = Year(Date) & m & d & "_US_Pack"
        USPickSheet = Year(Date) & m & d & "_US_Pick"
        SZPacksheet = Year(Date) & m & d & "_SZ_Pack"
        SZPickSheet = Year(Date) & m & d & "_SZ_Pick"
        SZ4PXSheet = Year(Date) & m & d & "_SZ_4PX"

        'MsgBox ResultSheet & vbCrLf & USPackSheet & vbCrLf & USPickSheet & vbCrLf & SZPacksheet & vbCrLf & SZPickSheet & vbCrLf & SZ4PXSheet

End Sub
Sub delete_today_sheet()

    Dim ifdelete As Integer
    ifdelete = MsgBox("没关系，删了还可以再算：）", vbOKCancel, "删还是不删？")
    If ifdelete = 1 Then
        Call set_result_sheets_name
        Application.DisplayAlerts = False
        If SheetExists(ResultSheet) = True Then
             sheets(ResultSheet).Delete
        End If
        If SheetExists(USPackSheet) = True Then
            sheets(USPackSheet).Delete
        End If
        If SheetExists(USPickSheet) = True Then
            sheets(USPickSheet).Delete
        End If
        If SheetExists(SZPacksheet) = True Then
            sheets(SZPacksheet).Delete

        End If
        If SheetExists(SZPickSheet) = True Then
            sheets(SZPickSheet).Delete
        End If
        If SheetExists(SZ4PXSheet) = True Then
            sheets(SZ4PXSheet).Delete
        End If
         Application.DisplayAlerts = True
    End If
End Sub
Sub insert_col_for_address()
    sheets(ResultSheet).Activate
        Columns("C:C").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("C1").Select
        ActiveCell.FormulaR1C1 = "address"

End Sub
Sub add_product_info()
    Dim k As Integer
    For k = 2 To Ordercount Step 1
        '-----将该行需要参与计算的单元格的值赋值给全局变量----
        Sale_number = sheets(ResultSheet).Cells(k, 1).Value
        Product_SKU = sheets(ResultSheet).Cells(k, 14).Value

        If Product_SKU <> "" Then
            Call get_product_info(Product_SKU, k)
        Else
            sheets(ResultSheet).Cells(k, 17).Value = "no sku"
            sheets(ResultSheet).Cells(k, 18).Value = "no sku"
        End If
    Next k
End Sub
Sub copy_sheet(ByVal souresheet As String, ByVal dessheet As String)
    '-----copy source sheet to des sheet and move des sheet to the position after source sheet-----
    sheets(souresheet).Copy after:=sheets(sheets.count)
    If SheetExists(dessheet) = True Then
        Application.DisplayAlerts = False
        sheets(dessheet).Delete
        Application.DisplayAlerts = True
    End If
    sheets(sheets.count).Name = dessheet
    'Sheets(Sheets.count).Move after:=Worksheets(souresheet)

End Sub
Sub formatpacksheet(ByVal sheetname As String)
    sheets(sheetname).Activate
    Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
End Sub
Sub formatresultsheet()
    sheets(ResultSheet).Activate
    Call addborderline(ResultSheet)
    Call freezefirstrow
    Columns("A:A").ColumnWidth = 6.13
    Columns("C:C").ColumnWidth = 23
    Columns("D:D").ColumnWidth = 13
    Columns("E:E").ColumnWidth = 16
    Columns("I:I").ColumnWidth = 6
    Columns("j:j").ColumnWidth = 5
    Columns("k:k").ColumnWidth = 7.38
    Columns("l:l").ColumnWidth = 8.13
    Columns("M:M").ColumnWidth = 12
    Columns("n:n").ColumnWidth = 36.75
    Columns("O:O").ColumnWidth = 9
    Columns("P:P").ColumnWidth = 5
    Columns("T:T").ColumnWidth = 12
    Columns("Z:Z").ColumnWidth = 8
    Columns("AH:AH").ColumnWidth = 13
    Columns("AI:AI").ColumnWidth = 11.75
    Cells.Select
    With Selection.font
        .Name = "宋体"
        .Size = 10
    End With
    With Selection
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        Rows.AutoFit
    End With
    Columns("M:M").Select
    Selection.NumberFormatLocal = "0_ "
    'hide useless columns
    Columns("F:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("V:Y").Select
    Selection.EntireColumn.Hidden = True
    Columns("AA:AI").Select
    Selection.EntireColumn.Hidden = True
    Columns("AK:AN").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Select
End Sub
Sub addebayaccount(ByVal sheetname As String)
    Dim i As Integer
    'MsgBox Ordercount & vbCrLf & Ebay_account
    For i = 2 To Ordercount
        sheets(sheetname).Cells(i, 1).Value = Ebay_account + CStr(sheets(sheetname).Cells(i, 1).Value)
    Next i
End Sub
Sub pre_process_resultsheet(ByVal sheetname As String)
    '--------in the original order sheet, if an order has multiple rows, only the first row has shipping service value,all next rows are null--------
    '--------this sub is to fill all null shipping service and delete the first row------
    sheets(sheetname).Activate
    '----------delete 2 empty rows on top of the sheets----------
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Ordercount = Ordercount - 2
    'MsgBox Ordercount
    '-------Hardcode here to define an array.
    Dim deleterow(1 To 1000) As Integer
    Dim u, j, iffirst As Integer
    Dim rowrange As String
    u = 1
    iffirst = 1

    For k = 2 To Ordercount Step 1
        If sheets(ResultSheet).Cells(k, 1).Value <> sheets(ResultSheet).Cells(k + 1, 1).Value Then
            iffirst = 1
        End If
        If sheets(ResultSheet).Cells(k, 1).Value = sheets(ResultSheet).Cells(k + 1, 1).Value And _
        iffirst = 1 Then
            For j = 3 To 11 Step 1
                sheets(ResultSheet).Cells(k + 1, j).Value = sheets(ResultSheet).Cells(k, j).Value
            Next j
            sheets(ResultSheet).Cells(k + 1, 34).Value = sheets(ResultSheet).Cells(k, 34).Value
            sheets(ResultSheet).Cells(k + 1, 35).Value = sheets(ResultSheet).Cells(k, 35).Value

            deleterow(u) = k
            u = u + 1
            iffirst = 0
        End If
    Next k

    '-------delete useless summary rows-----
    For j = 1 To u - 1
        rowrange = CStr(deleterow(j) - (j - 1)) & ":" & CStr(deleterow(j) - (j - 1))
        Rows(rowrange).Select
        Selection.Delete Shift:=xlUp
    Next j

    '------update Order count after delete rows
    Ordercount = Ordercount - u + 1
    'MsgBox u & vbCrLf & Ordercount

    '------delete 2 uesless rows at bottom of the sheet------
    rowrange = CStr(Ordercount + 3) & ":" & CStr(Ordercount + 3)
    Rows(rowrange).Select
    Selection.Delete Shift:=xlUp
    rowrange = CStr(Ordercount + 2) & ":" & CStr(Ordercount + 2)
    Rows(rowrange).Select
    Selection.Delete Shift:=xlUp

    '-----add lacked shipping service and country-------------
    Dim shippingservice As String
    'Dim nullitem As String
    For k = 2 To Ordercount Step 1
        If sheets(ResultSheet).Cells(k, 35).Value <> "" Then
            shippingservice = sheets(ResultSheet).Cells(k, 35).Value
        End If
        If sheets(ResultSheet).Cells(k, 35).Value = "" And _
        sheets(ResultSheet).Cells(k, 1).Value = sheets(ResultSheet).Cells(k - 1, 1).Value Then
            sheets(ResultSheet).Cells(k, 35).Value = shippingservice
        End If
    Next k
    For w = 3 To 11
        For k = 2 To Ordercount Step 1
'            'If sheets(ResultSheet).Cells(k, w).Value <> "" Then
'                nullitem = sheets(ResultSheet).Cells(k, w).Value
'            'End If
'            If sheets(ResultSheet).Cells(k, w).Value = "" And _
'            sheets(ResultSheet).Cells(k, 1).Value = sheets(ResultSheet).Cells(k - 1, 1).Value Then
'                sheets(ResultSheet).Cells(k, w).Value = nullitem
'            End If
            If sheets(ResultSheet).Cells(k, 1).Value = sheets(ResultSheet).Cells(k - 1, 1).Value Then
                sheets(ResultSheet).Cells(k, w).Value = sheets(ResultSheet).Cells(k - 1, w).Value
            End If
        Next k
    Next w
End Sub

Sub merge_same_orderid(ByVal sheetname As String, ByVal rowcount As Integer, ByVal sheettype As String)
    Dim start As Integer
    Dim idmergerange As String
    'Dim colnumber As Variant
    sheets(sheetname).Activate
    'Dim colarray() As Variant
    'colarray() = Array("A", "H", "I", "J", "K", "L")
    start = 0
    For k = 2 To rowcount Step 1
        If sheets(sheetname).Cells(k, 1).Value = sheets(sheetname).Cells(k + 1, 1).Value And _
        sheets(sheetname).Cells(k, 1).Value <> sheets(sheetname).Cells(k - 1, 1).Value Then
            start = k
        End If
        If (sheets(sheetname).Cells(k, 1).Value <> sheets(sheetname).Cells(k + 1, 1).Value) And (start <> 0) Then
            idmergerange = "A" & CStr(start) & ":A" & CStr(k)
            Call merge_cells(idmergerange)
            idmergerange = "H" & CStr(start) & ":H" & CStr(k)
            Call merge_cells(idmergerange)
            idmergerange = "I" & CStr(start) & ":I" & CStr(k)
            Call merge_cells(idmergerange)
            idmergerange = "J" & CStr(start) & ":J" & CStr(k)
            Call merge_cells(idmergerange)
            idmergerange = "K" & CStr(start) & ":K" & CStr(k)
            Call merge_cells(idmergerange)
            idmergerange = "L" & CStr(start) & ":L" & CStr(k)
            Call merge_cells(idmergerange)
            If sheettype = "US" Then
                idmergerange = "M" & CStr(start) & ":M" & CStr(k)
                Call merge_cells(idmergerange)
            ElseIf sheettype = "SZ" Then
                idmergerange = "B" & CStr(start) & ":B" & CStr(k)
                Call merge_cells(idmergerange)
                idmergerange = "C" & CStr(start) & ":C" & CStr(k)
                Call merge_cells(idmergerange)
                idmergerange = "D" & CStr(start) & ":D" & CStr(k)
                Call merge_cells(idmergerange)
            End If
            start = 0
        End If
    Next k
End Sub
Sub merge_cells(mergerange As String)
'
    Application.DisplayAlerts = False
    Range(mergerange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.merge
    Application.DisplayAlerts = True
End Sub
Sub freezefirstrow()
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

Sub mergeaddress()

    For k = 2 To Ordercount Step 1
        If InStr(1, sheets(ResultSheet).Cells(k, 36).Value, "USPS", vbTextCompare) <> 0 Then
        '------if sent to US, merge all info into one column--------
            If sheets(ResultSheet).Cells(k, 4).Value <> "" Then
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 4).Value
            End If
        End If

        '------merge address1 and address2 into 1 column regardless to US or other regions----

        If sheets(ResultSheet).Cells(k, 7).Value <> "" Then
            If sheets(ResultSheet).Cells(k, 3).Value <> "" Then
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 3).Value & Chr(10) & _
                    sheets(ResultSheet).Cells(k, 7).Value
            Else
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 7).Value
            End If
        End If
        If sheets(ResultSheet).Cells(k, 8).Value <> "" Then
            sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 3).Value & Chr(10) & _
                sheets(ResultSheet).Cells(k, 8).Value
        End If

        If InStr(1, sheets(ResultSheet).Cells(k, 36).Value, "USPS", vbTextCompare) <> 0 Then
        '------if sent to US, merge all info into one column--------
            If sheets(ResultSheet).Cells(k, 9).Value <> "" Then
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 3).Value & Chr(10) & _
                    sheets(ResultSheet).Cells(k, 9).Value
            End If
            If sheets(ResultSheet).Cells(k, 10).Value <> "" Then
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 3).Value & Chr(10) & _
                    sheets(ResultSheet).Cells(k, 10).Value
            End If
            If sheets(ResultSheet).Cells(k, 11).Value <> "" Then
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 3).Value & Chr(10) & _
                    sheets(ResultSheet).Cells(k, 11).Value
            End If
            If sheets(ResultSheet).Cells(k, 12).Value <> "" Then
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 3).Value & Chr(10) & _
                    sheets(ResultSheet).Cells(k, 12).Value
            End If
            If sheets(ResultSheet).Cells(k, 5).Value <> "" Then
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 3).Value & Chr(10) & _
                    sheets(ResultSheet).Cells(k, 5).Value
            End If
            If sheets(ResultSheet).Cells(k, 6).Value <> "" Then
                sheets(ResultSheet).Cells(k, 3).Value = sheets(ResultSheet).Cells(k, 3).Value & Chr(10) & _
                    sheets(ResultSheet).Cells(k, 6).Value
            End If
        End If
    Next k

End Sub
Sub result_sheet_sorting(ByVal sheetname As String, ByVal count As Integer, ByVal ordertype As Integer)
'
' ---------type 1时：按照shipping serviced倒序，按照order id正序-----------
' ---------type 0时：按照shipping service正序，按照order id正序-----------
' ---------type 2时：按照coutry 倒序，按照order id正序-----------
    sheets(sheetname).Activate
    Dim rowcount As Integer
    rowcount = count
    Dim keyrange, keyrange2, sortrange As String
    If ordertype = 1 Or ordertype = 0 Then
        keyrange = "AJ2:AJ" & CStr(rowcount)
    ElseIf ordertype = 2 Then
        keyrange = "L2:L" & CStr(rowcount)
    End If
    keyrange2 = "A2:A" & CStr(rowcount)
    sortrange = "A1:AP" & CStr(rowcount)
    'MsgBox keyrange & vbCrLf & sortrange

    If ordertype = 1 Then
        Worksheets(sheetname).Activate
        Cells.Select
        Worksheets(sheetname).sort.SortFields.Clear
        Worksheets(sheetname).sort.SortFields. _
            Add Key:=Range(keyrange), SortOn:=xlSortOnValues, Order:=xlDescending _
             , DataOption:=xlSortNormal
        Worksheets(sheetname).sort.SortFields. _
            Add Key:=Range(keyrange2), SortOn:=xlSortOnValues, Order:=xlAscending _
            , DataOption:=xlSortNormal
        With Worksheets(sheetname).sort
            .SetRange Range(sortrange)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    ElseIf ordertype = 0 Then
        Worksheets(sheetname).sort.SortFields. _
            Add Key:=Range(keyrange), SortOn:=xlSortOnValues, Order:=xlAscending _
             , DataOption:=xlSortNormal
        Worksheets(sheetname).sort.SortFields. _
            Add Key:=Range(keyrange2), SortOn:=xlSortOnValues, Order:=xlAscending _
            , DataOption:=xlSortNormal
        With Worksheets(sheetname).sort
            .SetRange Range(sortrange)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If

End Sub

Sub get_product_info(Product_SKU As String, row As Integer)
    '-------通过SKU找到产品的中文名称、采购专员并写入结果表
    For m = 2 To Productcount
        If sheets(ProductSheet).Cells(m, 1).Value = Product_SKU Then
            '-----中文名称------
            If sheets(ProductSheet).Cells(m, 21).Value = "" Then
                sheets(ResultSheet).Cells(row, 17).Value = "匹配为空"
            Else
                sheets(ResultSheet).Cells(row, 17).Value = sheets(ProductSheet).Cells(m, 21).Value
            End If

            '-----采购专员------
            If sheets(ProductSheet).Cells(m, 15).Value = "" Then
                sheets(ResultSheet).Cells(row, 18).Value = "匹配为空"
            Else
                sheets(ResultSheet).Cells(row, 18).Value = sheets(ProductSheet).Cells(m, 15).Value
            End If

            Exit Sub
        End If
    Next m
    sheets(ResultSheet).Cells(row, 17).Value = "无匹配"
    sheets(ResultSheet).Cells(row, 18).Value = "无匹配"

End Sub

Sub insert_column()
' insert columns into result sheet to contrain new data
'
    sheets(ResultSheet).Activate
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.SmallScroll Down:=-4

    'Add title for new added columns
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "中文名称"
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "采购员"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "图片"
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "快递单号"
    Range("T3").Select
End Sub
Sub addborderline(ByVal sheetname As String)
'
    sheets(sheetname).Select
    Cells.Select
    Range("F1").Activate
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

Function Getrowcount(ByVal sheetname As String) As Integer

    '-----------------计算表格中有多少行数据涉及计算------------------
    '------------根据第一列的个数计算-------------
    Dim i As Integer
    '-----------从第2行开始找不为空的单元格----------
    i = 4
    sheets(sheetname).Select
    Do While ActiveSheet.Cells(i, 1) <> ""
        i = i + 1
    Loop
    i = i - 1
    Getrowcount = i

End Function

Private Function SheetExists(sname) As Boolean
'如果活动工作簿中存在表SNAME则返回真
    Dim x As Object
    On Error Resume Next
    Set x = ActiveWorkbook.sheets(sname)
    If Err = 0 Then SheetExists = True _
        Else SheetExists = False
End Function
