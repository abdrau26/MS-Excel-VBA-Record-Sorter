- MS excel file is 002 SortingRecords.xlsm
- here we will sort data using VBA code.
- Excel sheet is `SORT RECORDS`.
- Columns are `Division` `Category` `Jan` `Feb` `Mar` `Total`
- The first step is to create user input box using VBA code 
- code is given below

```
Public Sub userInputForm()
Dim userInput As String
Dim promptMSG As String

    promptMSG = "Enter a numeric value to sort..." & vbCrLf & _
    "1-- Sort by Division" & vbCrLf & _
    "2-- Sort by Category" & vbCrLf & _
    "3-- Sort by Total"
    
    userInput = InputBox(promptMSG)

End Sub
```
- screenshot of input Box is given below

![[Screenshot 2026-05-13 174015.png]]

