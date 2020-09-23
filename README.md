<div align="center">

## Change Colour of Specified Text \(Like HTML tags\) in a RTF Box


</div>

### Description

This sifts through all the text in a RTF file and changes the colour of html tags. It can be used ewither automatically or on the fly (as you type into the RTF box). It currently Chnages the colour of html tags (defined as anything between "<" and ">") and html comments ("<!" to ">"). Enjoy! Please mail me with suggestions
 
### More Info
 
You need to have a RTF box with the default name (richtextbox1) on your form.

No side effects have been noticed.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eriond](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eriond.md)
**Level**          |Advanced
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eriond-change-colour-of-specified-text-like-html-tags-in-a-rtf-box__1-8939/archive/master.zip)

### API Declarations

```
You need to declare these variable within your forms declarations:
For the On the Fly Declare:
Dim previous as integer
For the automated declare:
posStart As Integer
Dim previousChar As String
```


### Source Code

```
On The Fly:
Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case 60
  'open tag
  RichTextBox1.SelLength = 0
  RichTextBox1.SelColor = &H8000000F
  previous = KeyAscii
 Case 62
  'close tag
  RichTextBox1.SelLength = 0
  RichTextBox1.SelText = ">"
  RichTextBox1.SelColor = &H0&
  previous = KeyAscii
  KeyAscii = 0
 Case 33
  'comments
  If previous = 60 Then
  RichTextBox1.SelStart = RichTextBox1.SelStart - 1
  RichTextBox1.SelLength = 1
  RichTextBox1.SelText = ""
  RichTextBox1.SelLength = 0
  RichTextBox1.SelColor = &HC00000
  RichTextBox1.SelText = "<!"
  previous = KeyAscii
  KeyAscii = 0
  End If
End Select
End Sub
Automated:
Sub ChangeColours()
 Dim posEnd As Integer
 i = 0
 For i = 0 To Len(RichTextBox1.Text)
  RichTextBox1.SelStart = i
  RichTextBox1.SelLength = 1
  If RichTextBox1.SelText = "<" Then 'start tag
   posStart = i
  End If
  If RichTextBox1.SelText = ">" Then 'end tag
   posEnd = i
  End If
  If RichTextBox1.SelText = "!" Then 'comment
   previousChar = "!"
  End If
  If posEnd <> 0 Then
   RichTextBox1.SelStart = posStart
   RichTextBox1.SelLength = posEnd - posStart + 1
   If previousChar <> "!" Then 'if not comment
    RichTextBox1.SelColor = &H8000000F
   Else:
    RichTextBox1.SelColor = &HC00000
    previousChar = " "
   End If
   RichTextBox1.SelStart = posStart + 1
   RichTextBox1.SelLength = 0
   RichTextBox1.SelColor = &H0&
   PosEnd = 0
   posStart = 0
  End If
  Next i
End Sub
```

