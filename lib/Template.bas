Attribute VB_Name = "Template"
Public serials, partouts, batchouts, kpi, text, mrp, code1, code2, code3, _
    assy_mrp, assy_serial, logs As String

Public cellA, cellB, cellC, cellD, cellE, cellF, cellG, cellH, cellI, cellJ, _
    cellK, cellL, cellM, cellN, cellO, cellP, cellQ As String

Sub excel(action As String)
    Excel_sheet.hook
    Application.ScreenUpdating = False
    
    'Clear Rows 2-3
    Worksheets("Shipping").chkStock.Visible = False
    Rows("2:3").ClearContents
    Rows("2:3").UnMerge
    Color.Reset
    
    'xlDiagonalDown ... xlInsideHorizontal = 5@12
    For itt = 5 To 12
        Rows("2:3").Borders(itt).LineStyle = xlNone
    Next
    Call Template.switch_case("GUI", action)
End Sub

Sub dropBox()
    With Worksheets("Shipping").cmbInput
        .Clear
        .ListFillRange = ""
        .AddItem "Close RMA"
        .AddItem "Mass Status Maintenance"
        .AddItem "Req"
        .AddItem "Req556"
        .AddItem "Swap"
        .AddItem "TECO"
        .AddItem "Create Material"
        .AddItem "Change Serial"
        .AddItem "Print, MB11 and TECO"
        .AddItem "IW72, outbound delivery"
        .ListIndex = 0
    End With
End Sub

Sub frame(y, x, width As Integer, cellText As String)
    Call data.write_cell(x, y, cellText)
    colN = Chr(x + 64) 'Convert number to letter
    Columns(colN).ColumnWidth = width
    
    'xlEdgeLeft ... xlEdgeRight = 7@10
    For itt = 7 To 10
        With objSheet.Cells(y, x).Borders(itt)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next
End Sub

Sub switch_case(step, action As String)
    Select Case action
        Case "Close RMA"
            Template.Close_RMA (step)
        Case "Mass Status Maintenance"
            Template.mass_status_maintenance (step)
        Case "Swap"
            Template.swap (step)
        Case "Req"
            Template.req (step)
        Case "Req556"
            Template.req556 (step)
        Case "TECO"
            Template.teco (step)
        Case "Create Material"
            Template.create_material (step)
        Case "Change Serial"
            Template.change_serial (step)
        Case "Print, MB11 and TECO"
            Template.mb11_teco (step)
        Case "IW72, outbound delivery"
            Template.iw72_out (step)
    End Select
End Sub

Sub Close_RMA(ByVal action As String)
    Select Case action
        Case "GUI"
            Range("G2:I2").Merge
            Call Template.frame(2, 7, 0, "Catalogue Code")
            Call Template.frame(3, 1, 21, "Serial")
            Call Template.frame(3, 2, 16, "PartOut")
            Call Template.frame(3, 3, 10, "BatchOut")
            Call Template.frame(3, 4, 4, "KPI")
            Call Template.frame(3, 5, 6, "Text")
            Call Template.frame(3, 6, 9, "MRP")
            Call Template.frame(3, 7, 10, "Symptome")
            Call Template.frame(3, 8, 10, "Défaut")
            Call Template.frame(3, 9, 10, "Assemblage")
            Call Template.frame(3, 10, 50, "Log")
        Case "read"
            serials = objSheet.Cells(i, 1).Value
            partouts = objSheet.Cells(i, 2).Value
            batchouts = objSheet.Cells(i, 3).Value
            kpi = objSheet.Cells(i, 4).Value
            text = objSheet.Cells(i, 5).Value
            mrp = objSheet.Cells(i, 6).Value
            code1 = objSheet.Cells(i, 7).Value
            code2 = objSheet.Cells(i, 8).Value
            code3 = objSheet.Cells(i, 9).Value
            logs = objSheet.Cells(i, 10).Value
        Case "execute"
            IW72.Full_run
            If IW72.partout = "" Then
                MB11.Config ("343")
                IW42.teco
            Else
                MB11.Config ("555")
                IQ08.partout
                MB11.Config ("501")
                IW42.teco
                VA02.partout
            End If
    End Select
End Sub

Sub mass_status_maintenance(ByVal action As String)
    Select Case action
        Case "GUI"
            Range("B2:G2").Merge
            Range("H2:P2").Merge
            Call Template.frame(2, 2, 0, "Put an 'x' to activate status")
            Call Template.frame(2, 8, 0, "1) a = activate" & vbLf & "2) r = remove" & vbLf & "3) blank = leave as is")
            Call Template.frame(3, 1, 21, "Serial")
            Call Template.frame(3, 2, 6, "TOEV")
            Call Template.frame(3, 3, 6, "EVAL")
            Call Template.frame(3, 4, 6, "HOLD")
            Call Template.frame(3, 5, 6, "REPA")
            Call Template.frame(3, 6, 6, "ESCL")
            Call Template.frame(3, 7, 6, "OTV")
            Call Template.frame(3, 8, 6, "BO")
            Call Template.frame(3, 9, 6, "ENG")
            Call Template.frame(3, 10, 6, "FA")
            Call Template.frame(3, 11, 6, "NPF")
            Call Template.frame(3, 12, 6, "PO")
            Call Template.frame(3, 13, 6, "PRD")
            Call Template.frame(3, 14, 6, "SCRP")
            Call Template.frame(3, 15, 6, "SWAP")
            Call Template.frame(3, 16, 6, "TS")
            Call Template.frame(3, 17, 50, "RMA long text(If needed)")
        Case "read"
            serials = objSheet.Cells(i, 1).Value
            cellB = objSheet.Cells(i, 2).Value
            cellC = objSheet.Cells(i, 3).Value
            cellD = objSheet.Cells(i, 4).Value
            cellE = objSheet.Cells(i, 5).Value
            cellF = objSheet.Cells(i, 6).Value
            cellG = objSheet.Cells(i, 7).Value
            cellH = objSheet.Cells(i, 8).Value
            cellI = objSheet.Cells(i, 9).Value
            cellJ = objSheet.Cells(i, 10).Value
            cellK = objSheet.Cells(i, 11).Value
            cellL = objSheet.Cells(i, 12).Value
            cellM = objSheet.Cells(i, 13).Value
            cellN = objSheet.Cells(i, 14).Value
            cellO = objSheet.Cells(i, 15).Value
            cellP = objSheet.Cells(i, 16).Value
            logs = objSheet.Cells(i, 17).Value
        Case "execute"
            IW72.Enter
            IW72.Status
    End Select
End Sub

Sub req(ByVal action As String)
    Select Case action
        Case "GUI"
            Call Template.frame(3, 1, 21, "Serial")
            Call Template.frame(3, 2, 50, "Commentaire")
        Case "read"
            serials = objSheet.Cells(i, 1).Value
            logs = objSheet.Cells(i, 2).Value
        Case "execute"
            IW72.req_in_out
            If in_out = "in" Then
                Misc.Zkciresrep
            ElseIf sloc <> "PL01" Then
                Call MB11.Config("555")
                Misc.Zkciresrep
            End If
    End Select
End Sub

Sub req556(ByVal action As String)
    Select Case action
        Case "GUI"
            Call Template.frame(3, 1, 21, "Serial")
            Call Template.frame(3, 2, 50, "Commentaire")
        Case "read"
            serials = objSheet.Cells(i, 1).Value
            logs = objSheet.Cells(i, 2).Value
        Case "execute"
            IW72.Enter
            IW72.Read_only
            Call MB11.Config("556")
    End Select
End Sub

Sub swap(ByVal action As String)
    Select Case action
        Case "GUI"
            Worksheets("Shipping").chkStock.Visible = True
            Template.Close_RMA ("GUI")
        Case "read"
            Template.Close_RMA ("read")
        Case "execute"
            IW72.Full_run
            Call VA02.swap
            MB11.Config ("555")
            IW42.teco
    End Select
End Sub

Sub teco(ByVal action As String)
    Select Case action
        Case "GUI"
            Call Template.frame(3, 1, 21, "Serial")
        Case "read"
            serials = objSheet.Cells(i, 1).Value
        Case "execute"
            IW72.Read_only
            IW42.teco
    End Select
End Sub

Sub mb11_teco(ByVal action As String)
    Select Case action
        Case "GUI"
            Call Template.frame(3, 1, 21, "Serial")
        Case "read"
            serials = objSheet.Cells(i, 1).Value
        Case "execute"
            IW72.Read_only
            IW72.Goto_object
            IW72.Printer
            If IW72.partout = "" Then
                MB11.Config ("343")
                IW42.teco
            Else
                MB11.Config ("555")
                IQ08.partout
                MB11.Config ("501")
                IW42.teco
                VA02.partout
            End If
    End Select
End Sub

Sub create_material(ByVal action As String)
    Select Case action
        Case "GUI"
            Range("D2:E2").Merge
            Call Template.frame(2, 4, 0, "Optional")
            Call Template.frame(3, 1, 21, "BLANK")
            Call Template.frame(3, 2, 21, "Assy Serial")
            Call Template.frame(3, 3, 10, "MRP")
            Call Template.frame(3, 4, 15, "Manuf Name")
            Call Template.frame(3, 5, 15, "Manuf Part")
            Call Template.frame(3, 6, 17, "Output")
        Case "read"
            serials = objSheet.Cells(i, 2).Value
            assy_mrp = objSheet.Cells(i, 3).Value
            manuf_name = objSheet.Cells(i, 4).Value
            manuf_part = objSheet.Cells(i, 5).Value
        Case "execute"
            IQ01.create
    End Select
End Sub

Sub change_serial(ByVal action As String)
    Select Case action
        Case "GUI"
            Call Template.frame(3, 1, 21, "Serial")
            Call Template.frame(3, 2, 17, "Assy Serial")
        Case "read"
            serials = objSheet.Cells(i, 1).Value
            assy_serial = objSheet.Cells(i, 2).Value
        Case "execute"
            IQ08.change
    End Select
End Sub

Sub iw72_out(ByVal action As String)
    Select Case action
        Case "GUI"
            Template.Close_RMA ("GUI")
        Case "read"
            Template.Close_RMA ("read")
        Case "execute"
            IW72.Full_run
            IW42.teco
            VA02.remove_block
    End Select
End Sub
