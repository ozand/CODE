Attribute VB_Name = "satructutal_discount"

Sub satructutal_discount_get_program()
Dim patch as String, shFile as String, shPrgrmIn as String

qCalculate = CInt(InputBox("number of quarter (program)"))
nm_ActWb = ActiveWorkbook.Name
myLib.VBA_Start

shFile = "PRGRM"
patch = "p:\DPP\Business development\Book commercial\DPP\#Расчётные\1. MONTHLY TASKS\phasage in_out_2016.xlsm"
wbSISOST = myLib.OpenFile(patch, shFile)    
Workbooks(wbSISOST).Activat
Sheets(shFile).Select

Dim prgs As Programs, prg As program
Set prgs = New programs
prgs.FillFromSheet ActiveSheet
Workbooks(wbSISOST).Close
Workbooks(nm_ActWb).Activate 
shPrgrmIn = shFile
myLib.CreateSh (shPrgrmIn)
myLib.sheetActivateCleer (shPrgrmIn)

i = 1
For Each prg In prgs

i = i + 1
    With prg

        n = 1: Cells(i, n) = .month
        n = n + 1: Cells(i, n) = .cd_partners
        n = n + 1: Cells(i, n) = .brand
        n = n + 1: Cells(i, n) = .type_vl
        n = n + 1: Cells(i, n) = .s_group
        n = n + 1: Cells(i, n) = .val
    End With

Next


myLib.VBA_End
End Sub 