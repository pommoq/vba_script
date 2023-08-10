Public Sub SetDataFieldsNumFormat()
    ' set Number Format Pivot table
    ' ตาม Option table 

    'สร้าง table tb_Nformat ไว้ที่ sheet
    'มีบางส่วน|fommat|width
    'default|#,000;[Red]-#,000;|20
    '%|0%;[Red]-0%;    6
    'sum|#.00,,;[Red]-#.00,,;|10

    nFormat = [tb_Nformat]
    Default_Format = nFormat(1, 2)
    Default_Width = nFormat(1, 3)


    Dim xPF As PivotField
    Dim WorkRng As Range
    Set WorkRng = Application.Selection
    With WorkRng.PivotTable
       .ManualUpdate = True
       For Each xPF In .DataFields
          With xPF
            '.Function = xlSum
            cFormat = Default_Format
            cWidth = Default_Width
            cNum = xPF.DataRange.Column
            For i = 2 To UBound(nFormat)
                If InStr(xPF.Name, nFormat(i, 1)) > 0 Then
                    cFormat = nFormat(i, 2)
                    cWidth = nFormat(i, 3)
                End If
            Next
            .NumberFormat = "#.00;[Red]-#.00;"
            .NumberFormat = cFormat
            If cWidth = "" Then
                Columns(cNum).EntireColumn.AutoFit
            Else
                Columns(cNum).EntireColumn.ColumnWidth = cWidth
            End If
          End With
       Next
       .ManualUpdate = False
    End With
End Sub
