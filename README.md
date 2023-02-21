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
