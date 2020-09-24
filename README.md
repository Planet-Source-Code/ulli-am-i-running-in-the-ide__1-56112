<div align="center">

## Am I running in the IDE?


</div>

### Description

The proper way to find out whether your code is running in the IDE or was compiled
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ULLI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ulli.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ulli-am-i-running-in-the-ide__1-56112/archive/master.zip)





### Source Code

```
'this code goes into a class named cEnvironment
Option Explicit
Public Enum eEnvironment
  EnvironIDE = 1
  EnvironCompiled = 2
End Enum
Public Property Get QueryEnvironment() As eEnvironment
  QueryEnvironment = EnvironCompiled
  Debug.Assert Not SetToIDE(QueryEnvironment)
End Property
Private Function SetToIDE(Env As eEnvironment) As Boolean
  Env = EnvironIDE
End Function
'make QueryEnvironment the default property of class cEnvironment
'------------------------------------------
'and then use this anywhere in your code
Private Sub Something()
 Dim Environment As New cEnvironment
  Print IIf(Environment = EnvironIDE, " I am running in the IDE", _
                    " Somebody had mercy and compiled me")
'  you don't normally print the result, so you might type this...
'  If Environment =
'  ..and you will see the two possibilities in VB's popup
End Sub
```

