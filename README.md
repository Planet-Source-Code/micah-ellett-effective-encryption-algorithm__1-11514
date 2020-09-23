<div align="center">

## Effective Encryption Algorithm


</div>

### Description

Encrypts and decrypts strings
 
### More Info
 
Simply:

DMEncrypt "Text"

and

DMDecrypt "Text"

Pretty straightforward

2 functions return a string value, either encrypted or decrypted text.

None that I know of


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Micah Ellett](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/micah-ellett.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/micah-ellett-effective-encryption-algorithm__1-11514/archive/master.zip)





### Source Code

```
'THIS FUNCTION ENCRYPTS THE INPUT
Public Function DMEncrypt(strText As String)
On Error GoTo Xit
Dim Combine As String, i As Integer, Temp As String
Combine = ""
Temp = ""
For i = 1 To Len(strText) - 1 Step 2
  If Len(Trim(Str(Asc(Mid(strText, i, 1))))) < 3 Then
    Temp = "0" & Trim(Str(Asc(Mid(strText, i, 1))))
  Else
    Temp = Trim(Str(Asc(Mid(strText, i, 1))))
  End If
  Combine = Combine & Temp
  If Len(Trim(Str(Asc(Mid(strText, i + 1, 1))))) < 3 Then
    Temp = "0" & Trim(Str(Asc(Mid(strText, i + 1, 1))))
  Else
    Temp = Trim(Str(Asc(Mid(strText, i + 1, 1))))
  End If
  Combine = Combine & Temp
Next i
Temp = ""
For i = 1 To Len(Combine)
  Temp = Temp & Chr(Asc(Mid(Combine, i, 1)) + 128)
Next i
DMEncrypt = Temp
Clipboard.SetText Temp
Exit Function
Xit:
DMEncrypt = "{{ Error encrypting }}"
Exit Function
End Function
'THIS FUNCTION DECRYPTS THE INPUT
Public Function DMDecrypt(strText As String)
On Error GoTo Xit
Dim Combine As String, i As Integer, Temp As String, Temp2 As Integer
Combine = ""
For i = 1 To Len(strText)
  Combine = Combine & Chr(Asc(Mid(strText, i, 1)) - 128)
Next i
Temp = ""
For i = 1 To Len(Combine) Step 3
  Temp2 = Mid(Combine, i, 3)
  Temp = Temp & Chr(Temp2)
Next i
DMDecrypt = Temp
Exit Function
Xit:
DMDecrypt = "{{ Error encrypting }}"
Exit Function
End Function
```

