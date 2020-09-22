<div align="center">

## Is in Array Function


</div>

### Description

Find if a value exists in an array WITHOUT LOOPING. Often we need to find out if a value exists in an array. This one does it VERY FAST.

NOTE: This Function only return True / False regarding the Existence of a Value. If you need the Index you will have to LOOP.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Gillham](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-gillham.md)
**Level**          |Intermediate
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-gillham-is-in-array-function__1-25714/archive/master.zip)





### Source Code

```
Public Function IsInArray(FindValue As Variant, arrSearch As Variant) As Boolean
 On Error GoTo LocalError
 If Not IsArray(arrSearch) Then Exit Function
 If Not IsNumeric(FindValue) Then FindValue = UCase(FindValue)
 IsInArray = InStr(1, vbNullChar & Join(arrSearch, vbNullChar) & vbNullChar, vbNullChar & FindValue & vbNullChar) > 0
Exit Function
LocalError:
 'Justin (just in case)
End Function
```

