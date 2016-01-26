####Module : MExcelHelper  
#####Type : Module  
#####Description : Functions to work with the grid and other aspects of Excel  
***  
Option Explicit  
Option Private Module  
***  
Public Function ColumnNumbertoLetter(ByVal intColumn As Integer) As String  
***  
>Description : Converts the column number to the column letter  
>Parameters  :  
> > 1. intColumn : the number to convert to a column letter  
>  
>Returns : the column letter given the number  
>Dependencies: none  
>Notes: none  
>Usage :  
```  
Dim s as string  
s = ColumnNumbertoLetter(4)  
Debug.Assert(s = "d")  
```  
End Function  
***  
Public Sub ResetAppSettings()  
***  
>Description : Resets the application environment to the usual settings  
>Parameters : none  
>Returns : none  
>Dependencies: none  
>Notes: none  
>Usage :  
```  
ResetAppSettings  
```  
End Sub  
