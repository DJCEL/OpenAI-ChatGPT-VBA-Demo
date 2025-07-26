Attribute VB_Name = "ThisWorkbook"
Private Sub Workbook_Open()

    Dim WS As Excel.Worksheet
    Dim WS_params As Excel.Worksheet
    Dim strInput As String
    Dim question As String
    Dim MODEL As String
    Dim sourceRange As Excel.Range
    Dim cb As ComboBox
    Dim strOpenAI_Model As String

    Set WS = ThisWorkbook.Sheets("VBA")
    Set WS_params = ThisWorkbook.Sheets("params")
    
    
    Set cb = ThisWorkbook.Sheets("VBA").ComboBox1
    Set sourceRange = WS_params.Range("A2:A33")
    cb.List = sourceRange.Value
    
    ThisWorkbook.Sheets("VBA").ComboBox1 = "gpt-4.1"
    
    'Other fields
    ThisWorkbook.Sheets("VBA").TextBox1 = ""
    ThisWorkbook.Sheets("VBA").TextBox2 = ""
    ThisWorkbook.Sheets("VBA").TextBox3 = ""
    ThisWorkbook.Sheets("VBA").TextBox4 = ""
    ThisWorkbook.Sheets("VBA").TextBox5 = ""
    ThisWorkbook.Sheets("VBA").TextBox6 = ""
    ThisWorkbook.Sheets("VBA").TextBox7 = ""
    ThisWorkbook.Sheets("VBA").TextBox9 = ""
    ThisWorkbook.Sheets("VBA").TextBox10 = ""
    ThisWorkbook.Sheets("VBA").TextBox11 = ""
    
    strOpenAI_Model = ThisWorkbook.Sheets("VBA").ComboBox1
    strInput = ThisWorkbook.Sheets("VBA").TextBox2
    
    question = OpenAI_InputRole2JSON("user", strInput)
    MODEL = OpenAI_Model2JSON(strInput, strOpenAI_Model)
    
    ThisWorkbook.Sheets("VBA").TextBox6 = question
    ThisWorkbook.Sheets("VBA").TextBox7 = MODEL

End Sub
