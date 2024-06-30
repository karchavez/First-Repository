Attribute VB_Name = "Module1"
Option Explicit

Sub ElemOp()
Dim file As String, Folder As String, Status, rCode, rTotal, rpmPos, rpmPos2 As String
Dim Cntfile As Variant
Dim i As Long, c As Long, startRow As Long, r As Long
Dim srw As Byte, frw As Byte
Dim thisWB As Object, Code As Object, Total As Object, Bundle As Object, Insp As Object

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
'ActiveSheet.DisplayPageBreaks = False

Set thisWB = ThisWorkbook.Worksheets("Original Data")
Folder = thisWB.Range("c1").Value & "\"


'=====CONTAR ARCHIVOS========
Cntfile = Dir(Folder)
   While (Cntfile <> "")
     Cntfile = Dir
     c = c + 1
   Wend
file = Dir(Folder & "*.xls")
'=============================

    Do While file <> ""
    Workbooks.Open Folder & file, UpdateLinks:=False
            
        With Workbooks(file)

'Rangos de Busqueda Columna Code y Total:
rCode = "C:C"
rTotal = "I:I"

'Rangos datos de formula de costura
rpmPos = "C1:C300"
rpmPos2 = "B10:B600"


'=============================



 
    

    

              Workbooks(file).Sheets(1).Select 'Desagrupa las hojas
              For i = 1 To Workbooks(file).Sheets.Count   'Revisa cada tab del archivo
              
              If Workbooks(file).Sheets(i).Visible = xlSheetVisible Then 'Solo toma hojas visibles
              
                         startRow = thisWB.Range("A1048576").End(xlUp).Row + 1 'Fila donde inicia a copiar en este archivo
                         Set Code = Workbooks(file).Sheets(i).Range(rCode).Find("Code")
                         Set Total = Workbooks(file).Sheets(i).Range(rTotal).Find("Total")
                         Set Bundle = Workbooks(file).Sheets(i).Range("F:F").Find("bundle TMU") 'Cambiar a como esta escrito en el archivo****
                         Set Insp = Workbooks(file).Sheets(i).Range("F:F").Find("inspection") 'Cambiar a como esta escrito en el archivo****
                        
                         
              On Error Resume Next
                         
                         srw = Code.Row + 1
                         frw = Total.End(xlUp).Row
           
                    If srw <> 0 Then
                                
                        Workbooks(file).Sheets(i).Range("B" & srw, "D" & frw).Copy 'Copia secuencia, codigo y elementos
                        thisWB.Range("I" & startRow).PasteSpecial xlValues
                        Workbooks(file).Sheets(i).Range("H" & srw, "J" & frw).Copy
                        thisWB.Range("L" & startRow).PasteSpecial xlValues 'Copia frecuencia, tmu y total tmu
                        Workbooks(file).Sheets(i).Range("G" & srw, "G" & frw).Copy
                        thisWB.Range("S" & startRow).PasteSpecial xlPasteValuesAndNumberFormats 'Copia distancia de costura
                       
                        '--------------------------------------------------------------------------------------------------------------
                    
                        Workbooks(file).Sheets(i).Range("G" & Bundle.Row).Copy      'Copia Bundle
                        thisWB.Range("M" & thisWB.Range("J1048576").End(xlUp).Row + 1).PasteSpecial xlValues
                        thisWB.Range("J" & thisWB.Range("J1048576").End(xlUp).Row + 1, "K" & thisWB.Range("J1048576").End(xlUp).Row + 1) = "Bundle"
                        thisWB.Range("L" & thisWB.Range("J1048576").End(xlUp).Row) = "1"
'                        '--------------------------------------------------------------------------------------------------------------
                        Workbooks(file).Sheets(i).Range("G" & Insp.Row).Copy      'Copia Insp
                        thisWB.Range("M" & thisWB.Range("J1048576").End(xlUp).Row + 1).PasteSpecial xlValues
                        thisWB.Range("J" & thisWB.Range("J1048576").End(xlUp).Row + 1, "K" & thisWB.Range("J1048576").End(xlUp).Row + 1) = "Inspection"
                        thisWB.Range("L" & thisWB.Range("J1048576").End(xlUp).Row) = "1"
                        '--------------------------------------------------------------------------------------------------------------
                       
                        
                        r = thisWB.Range("J1048576").End(xlUp).Row
                        
                        'Buscar Datos Formula de Costura
                        
                        With thisWB
                        
                            .Range("A" & startRow, "A" & r) = Workbooks(file).Name
                            .Range("B" & startRow, "B" & r) = Workbooks(file).Sheets(i).Name  'Nombre de tab
                            .Range("C" & startRow, "C" & r) = Workbooks(file).Sheets(i).Range("C6")                                         'Plant
                            .Range("E" & startRow, "E" & r) = Workbooks(file).Sheets(i).Range("C4")                                         'Description
                            .Range("F" & startRow, "F" & r) = Workbooks(file).Sheets(i).Range("I6")                                         'Measure Range
                            .Range("G" & startRow, "G" & r) = Workbooks(file).Sheets(i).Range("D" & WorksheetFunction.Match(" RPM", Workbooks(file).Sheets(i).Range(rpmPos), 0)).Value 'RPM
                            .Range("H" & startRow, "H" & r) = Workbooks(file).Sheets(i).Range("D" & WorksheetFunction.Match(" RPM", Workbooks(file).Sheets(i).Range(rpmPos), 0) + 1).Value 'SPI
                            .Range("P" & startRow, "P" & r) = Workbooks(file).Sheets(i).Range("D" & WorksheetFunction.Match(" RPM", Workbooks(file).Sheets(i).Range(rpmPos), 0) + 3).Value 'GTF Value'
                            .Range("Q" & startRow, "Q" & r) = Workbooks(file).Sheets(i).Range("D" & WorksheetFunction.Match("RPM", Workbooks(file).Sheets(i).Range(rpmPos2), 0) + 11).Value 'GTF TMU
                            .Range("R" & startRow, "R" & r) = Workbooks(file).Sheets(i).Range("F" & WorksheetFunction.Match("RPM", Workbooks(file).Sheets(i).Range(rpmPos2), 0) + 11).Value 'HSF
                            .Range("T" & startRow, "T" & r) = Workbooks(file).Sheets(i).Range("F" & WorksheetFunction.Match(" RPM", Workbooks(file).Sheets(i).Range(rpmPos), 0) + 1).Value 'LOS (Runoff)
                            .Range("U" & startRow, "U" & r) = Workbooks(file).Sheets(i).Range("G" & WorksheetFunction.Match("RPM", Workbooks(file).Sheets(i).Range(rpmPos2), 0) + 11).Value 'SS
                            .Range("V" & startRow, "V" & r) = Workbooks(file).Sheets(i).Range("F" & WorksheetFunction.Match(" RPM", Workbooks(file).Sheets(i).Range(rpmPos), 0) + 3).Value 'P Value
                            .Range("W" & startRow, "W" & r) = Workbooks(file).Sheets(i).Range("H" & WorksheetFunction.Match("RPM", Workbooks(file).Sheets(i).Range(rpmPos2), 0) + 11).Value 'P TMU
                            .Range("X" & startRow, "X" & r) = Workbooks(file).Sheets(i).Range("D" & WorksheetFunction.Match(" RPM", Workbooks(file).Sheets(i).Range(rpmPos), 0) + 2).Value 'PFD
'                            .Range("Y" & startRow, "Y" & r) = Workbooks(file).Sheets(i).Range("C" & WorksheetFunction.Match(" RPM", Workbooks(file).Sheets(i).Range(rpmPos), 0) + 4).Value 'Factor Solo para Delta Cortes
                            .Range("AD" & startRow, "AD" & r) = Workbooks(file).Sheets(i).Range("C7")  'Machine
                            
                        End With
                        
                    Else: End If
                                           
                    srw = 0
                    Bundle.Row = 0
                   
                    
              End If
              Next i     'Siguiente pestana

                                                
Application.Wait (Now + TimeValue("00:00:05"))
Application.CutCopyMode = False
Workbooks(file).Close SaveChanges:=False


        End With    'Siguiente Archivo
                                      
file = Dir
Application.Wait (Now + TimeValue("00:00:05"))
Loop
            

' Call Blanks
' Call Formato
 thisWB.Activate
 Range("B3").Select


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True


End Sub

Sub Formato()

    Range("L1").Copy
    Range("O4").PasteSpecial xlPasteFormulas
    Range("O4").Select
    Selection.AutoFill Destination:=Range("O4", "O" & Range("A1048576").End(xlUp).Row)
    Range("O4", "O" & Range("B1048576").End(xlUp).Row).Copy
    Range("L4", "L" & Range("B1048576").End(xlUp).Row).PasteSpecial xlPasteValuesAndNumberFormats
    Range("O4", "O" & Range("B1048576").End(xlUp).Row).ClearContents


End Sub

Sub Blanks()
 On Error GoTo sig
Range("J4", "J" & Range("B1048576").End(xlUp).Row).Select
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete
sig:

End Sub


