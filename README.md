<div align="center">

## Convert StringIP to NumericIP


</div>

### Description

Converts a string ip address ("192.168.0.1") to a Long number (3232235521). One of the resons to do this would be to store IP addresses in databases. Numbers greatly reduce the size required to store this information.
 
### More Info
 
asNewIP - String IP address ("192.168.0.1") to convert to a number.

This function assumes that your IP address has 4 integers delimited by decimals and that the numbers range from 0 to 255.

Returns a Long Integer representing the IP address (3232235521)

If storing number within an Access Database, The variable type "Long" ranges from -2147483648 to 2147483647. You will need to subtract 2147483648 from the number before querying the database.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-convert-stringip-to-numericip__4-6265/archive/master.zip)

### API Declarations

Copyright (C) 1999, Lewis Moten. All rights reserved.


### Source Code

```
Function CLngIP(ByVal asNewIP)
	Dim lnResults
	Dim lnIndex
	Dim lnIpAry
	' Split the IP address using the dot as a delimiter
	lnIpAry = Split(asNewIP, ".", 4)
	' Loop through each number in the IP address
	For lnIndex = 0 To 3
		' If we are not working with the last number...
		If Not lnIndex = 3 Then
			' Convert the number to a value range that can be parsed from the others
			lnIpAry(lnIndex) = lnIpAry(lnIndex) * (256 ^ (3 - lnIndex))
		End If
		' Add the number to the results
		lnResults = lnResults + lnIpAry(lnIndex)
	Next
	' If storing number within an Access Database,
	' The variable type "Long" ranges from -2147483648 to 2147483647
	' You will need to subtract 2147483648 from the number
	' before querying the database.
	' lnResults = lnResults - 2147483648
	' Return the results
	CLngIP = lnResults
End Function
```

