VBA  Code

Sub addnewdata()

Sheets("data").Activate
Range("a2").Select

Do Until ActiveCell.Value = ""

ActiveCell.Offset(1, 0).Select

Loop

ActiveCell.Value = Sheets("form").Range("CLIENT").Value


ActiveCell.Offset(0, 1).Select

ActiveCell.Value = Sheets("form").Range("PDATE").Value

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = Sheets("form").Range("AMOUNT").Value


Sheet1.Activate
Sheet1.Range("C14").Value = "Last Execution info: Data Submitted Suceesfully " & Now()

End Sub
