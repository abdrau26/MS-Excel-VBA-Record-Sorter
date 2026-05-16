- MS excel file is `002 SortingRecords.xlsm`
- here we will sort data using VBA code.
- the sorting VBA by respective column has been generated using `Record Macro`
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
![user input form](user-input-form.png)


- now we are adding `IF` `ELSEIF` conditions. the code sorts the table according to column number provided by user. the code is given below:

```

Public Sub userInputForm()


Dim userInput As String
Dim promptMSG As String

    promptMSG = "Enter a numeric value to sort..." & vbCrLf & _
    "1-- Sort by Division" & vbCrLf & _
    "2-- Sort by Category" & vbCrLf & _
    "3-- Sort by Total"
    
    userInput = InputBox(promptMSG)
    
    If userInput = "1" Then
    DivisionSort
    ElseIf userInput = "2" Then
    CategorySort
    ElseIf userInput = "3" Then
    TotalSort
    
    
End If




End Sub
```
- Now if the user input does not corresponds to columns, it shows an error message. the code in `ELSE` statement is given as:

```
 Else
    tryAgain = MsgBox("Invalid Value. Want to try again?", vbYesNo)
    If tryAgain = 6 Then
    userInputForm
```
