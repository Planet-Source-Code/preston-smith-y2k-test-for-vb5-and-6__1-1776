<div align="center">

## Y2K Test for VB5 and 6


</div>

### Description

This program test the Visual Basic program you are currently running. As you will see.
 
### More Info
 
You can adjust the dates for Jan 1st 1930, Jan 1st 2029, Jan 1st 00 and Feb 29 or 30 2000. To see if your version is compliant then you should use 1/1/00 for the year 2000, and 2/29/00 for the Leap Year date (this will cause a run-time error, if NOT compliant). I made a big mistake before, the date regarding the leap year, should not cause a run time error if it is compliant.

If you set up the code to check for the leap year date you should have a run-time error. If you do NOT have an error that is good. If you actually see a date then your version of VB is not compliant with Y2K. The updates and patches may be available at microsoft's MSDN websight. Problems? Get the SP3 updates, Go here (VB6): http://msdn.microsoft.com/vstudio/sp/default.asp Go here (VB5): http://msdn.microsoft.com/vstudio/sp/vs97/default.asp


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Preston Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/preston-smith.md)
**Level**          |Unknown
**User Rating**    |5.9 (606 globes from 102 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/preston-smith-y2k-test-for-vb5-and-6__1-1776/archive/master.zip)





### Source Code

```
Sub Form_Load()
Dim MyDate as Date
MyDate = "1/1/00" 'Or
'MyDate = "1/1/29" 'Returns 1/1/2029
'MyDate = "1/1/30" 'Returns 1/1/1930
'MyDate = "2/29/00" 'The Leap Year Date (Usually causes the most probs)
MsgBox Format(MyDate, "mm/dd/yyyy")
End Sub
```

