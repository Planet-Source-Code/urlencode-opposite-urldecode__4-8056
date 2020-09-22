<div align="center">

## URLEncode opposite \(URLDecode\)


</div>

### Description

The function Server.URLEncode converts a inputstring to an URL encoded outputstring. For example:

input: Server.URLEncode("part1?part2")

output: "part1%3Fpart2"

But what if you need the opposite functionality ? There is no function available for this so you have build this yourself. How ? By using regular expressions, of course.
 
### More Info
 
This function searches for %[HEX VALUE][HEX VALUE] and replaces them by converting [HEX VALUE][HEX VALUE] to an integer an converting the integer to an ASCII character.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\-\-\-\-\-\-\-\-\-\-](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/urlencode-opposite-urldecode__4-8056/archive/master.zip)





### Source Code

```
Function URLDecode(sText)
	sDecoded = sText
  Set oRegExpr = Server.CreateObject("VBScript.RegExp")
  oRegExpr.Pattern = "%[0-9,A-F]{2}"
  oRegExpr.Global = True
  Set oMatchCollection = oRegExpr.Execute(sText)
  For Each oMatch In oMatchCollection
		sDecoded = Replace(sDecoded,oMatch.value,Chr(CInt("&H" & Right(oMatch.Value,2))))
  Next
  URLDecode = sDecoded
End Function
```

