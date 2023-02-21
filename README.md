# DQ

Sub Get_Data_From_File2()
Dim FileToOpen As Variant
Dim OpenBook As Workbook
Application.ScreenUpdating = False
    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.xlsx*),*xlsx*")
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        
            Dim c As Range
            Dim firstAddress As String

            With Worksheets(1).Range("A1:L40")
            Set c = .Find("Item", LookIn:=xlValues)
            If Not c Is Nothing Then
            firstAddress = c.Address
            MsgBox (firstAddress)
            
        End If
    End With

    Call FindString
        OpenBook.Sheets(1).Range("A1:E20").Copy
        ThisWorkbook.Worksheets("Sheet1").Range("A10").PasteSpecial xlPasteValues
        OpenBook.Close False
    End If
    
  
Application.ScreenUpdating = True
    
End Sub

Sub FindString()

    Dim c As Range
    Dim firstAddress As String

    With Worksheets(1).Range("A1:L40")
        Set c = .Find("dash", LookIn:=xlValues)
        If Not c Is Nothing Then
            firstAddress = c.Address
            MsgBox (firstAddress)
            
        End If
    End With
    
    
    Sub define_cell_adress()

Dim ThisPos As Range
With Range("A1:J100")
    Set ThisPos = .Find(What:="dash", LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
    If Not ThisPos Is Nothing Then
        Cell_Add = Split(ThisPos.Address, "$")
        ThisRow = Cell_Add(1)
        ThisCol = Cell_Add(2)
        MsgBox (ThisRow)
        MsgBox (ThisCol)
    End If
End With
End Sub
