<div align="center">

## CSV Parser


</div>

### Description

'    This function is for use when parsing(splitting) a data string that

'    has a comma delimiter. The normal VB Split function does not take into

'    consideration of a comma embedded within a Fields' data string and

'    will parse the information incorrectly.

'

'    This function takes into consideration the a data field may contain

'    a comma and parses the data as entire string. The data string being defined

'    as the data between the two Double-Quote marks. This function also

'    prunes the leading and trailing double quote marks
 
### More Info
 
Data String to Parse.

This should be able to run under VB 5.0.

' Returns  : Single-Dimension array, same result that you get

'       from the SPLIT Function.

' Notes   : Does NOT Correct improperly formatted Numeric amounts that

'      : contain a comma for the thousands placement, unless the number has

'      : leading and trailing Double-Quote marks.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/csv-parser__1-11993/archive/master.zip)





### Source Code

```
Option Explicit
Public Function PSplit(sInstring As String) As Variant
'
' Author: Scott Bingham, July 2000
'
' Function Name: ' PSplit = Proper Split.
'
' Versions: VB 6.0 (Should work with 5.0 also)
'
' Overview:
'    This function is for use when parsing(splitting) a data string that
'    has a comma delimiter. The normal VB Split function does not take into
'    consideration of a comma embedded within a Fields' data string and
'    will parse the information incorrectly.
'
'    This function takes into consideration the a data field may contain
'    a comma and parses the data as entire string. The data string being defined
'    as the data between the two Double-Quote marks. This function also
'    prunes the leading and trailing double quote marks
'
' Notes   : Does NOT Correct improperly formatted Numeric amounts that
'      : contain a comma for the thousands placement, unless the number has
'      : leading and trailing Double-Quote marks.
'
' Errors  : NONE
'
' Call   : X() = PSplit(datastring to split.)
'
' Returns  : Single-Dimension array, same result that you get
'       from the SPLIT Function.
  Dim sDelim$, iStringLength%, iDelimPosition%, sDoubleQuoteMark$
  Dim iIndex%, aData1() As String, sDatafield$
  Dim iDQPos1%, iDQPos2%
  '
  sDoubleQuoteMark = Chr$(34)
  sDelim = ","
  iStringLength = Len(sInstring)
  iIndex = 0
  '
  ' if the length of the data string is greater than zero
  If iStringLength > 0 Then
    ' search for a sDelimiter in the datastring
    iDelimPosition = InStr(sInstring, sDelim)
    '
    Do While iDelimPosition <> 0
      ' do while there is a sDelimiter
      ' search for a quote-enclosure set.
      iDQPos1 = InStr(sInstring, sDoubleQuoteMark)
      sDatafield = ""
      '
      If iDQPos1 <> 0 And iDQPos1 < iDelimPosition Then
        ' found Double quote mark, and it is found BEFORE
        ' the sDelimiter. Search for matching Double Quote Mark
        iDQPos2 = InStr(iDQPos1 + 1, sInstring, sDoubleQuoteMark)
        If iDQPos2 <> 0 Then
          If iDQPos2 = Len(sInstring) Then
          ' this is the last field of data so we remove the
          ' surrounding Double-Quote Marks.
            sInstring = Right(sInstring, Len(sInstring) - 1)
            sInstring = Left(sInstring, Len(sInstring) - 1)
            'exit the Do loop and
            Exit Do
          End If
          ' Just found the Matching double Quote Mark
          ' data field ends at iDQPos2, not iDelimPosition
          sDatafield = Left(sInstring, iDQPos2)
          sInstring = Right(sInstring, Len(sInstring) - (Len(sDatafield) + 1))
          sDatafield = Right(sDatafield, Len(sDatafield) - 1)
          sDatafield = Left(sDatafield, Len(sDatafield) - 1)
          iIndex = iIndex + 1
        Else
          ' unmatched double quote usually specifies error with the
          ' data being read in.
        End If
      Else
        If iDQPos1 <> 0 Then
          ' Quote mark is FOUND AFTER the sDelimiter meaning the
          ' data to the sDelimiter is ok to use as a full field.
          ' Data ends at the sDelimiter.
          sDatafield = Left(sInstring, iDelimPosition - 1)
          sInstring = Right(sInstring, Len(sInstring) - (Len(sDatafield) + 1))
          iIndex = iIndex + 1
        Else
          ' there is NO double Quote Mark Found.
          sDatafield = Left(sInstring, iDelimPosition - 1)
          sInstring = Right(sInstring, Len(sInstring) - iDelimPosition)
          iIndex = iIndex + 1
        End If
      End If
      ReDim Preserve aData1(iIndex)
      aData1(iIndex) = sDatafield
      iDelimPosition = InStr(sInstring, sDelim)
    Loop
    iIndex = iIndex + 1
    ReDim Preserve aData1(iIndex)
    aData1(iIndex) = sInstring
  Else
  End If
  PSplit = aData1
End Function
```

