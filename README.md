# Excel-Macro-Scripts

You can create a VBA macro to automate link generation in your spreadsheet. 

If you have a directory on your website that contains a gallery of photos, you can enter the photo names into a spreadsheet and then
generate the urls via a macro.

ie: Column A contains:
photo1.jpg
photo2.jpg
photo3.jpg

The following function will append the parent address that these files exist in, underline, color blue, and center the cells:

```vba
Sub CreateHyperLinks()
    Dim range As Range
    Dim fileName As String

    'Set range = ActiveSheet.Range("A:A")
        For Each range In ActiveSheet.Range("A:A").Cells
        fileName = range.Value
            If IsNumeric(fileName) Then
                range.Parent.Hyperlinks.Add Anchor:=range, Address:="http://mywebsite.com/gallery/" & fileName, SubAddress:= _
                    "", TextToDisplay:=fileName
                With range.Font
                    .ColorIndex = xlAutomatic
                    .Underline = xlUnderlineStyleNone
                    .Color = vbBlueb
                End With
                With range.Characters().Font
                    .Underline = xlUnderlineStyleSingle
                    .Color = vbBlue
                End With
            End If
        Next
        
        Columns("A").HorizontalAlignment = xlCenter
End Sub
```
A Sharper Excel Interop Design

Understanding how to use Excel Interop for spreadsheet manipulation in C# without hours of googling.
How to add Excel Interop to your project

    Create your project
    In the Solution Explorer, right click on your project -> Add -> Reference...
    On the left-hand side, select the COM tab.
    In the "Search COM" box in the top right, search for "excel"
    One of your options should be "Microsoft Excel version_number Object Library". In my case version_number = 16.0.
    Check the box and press OK.

You have successfully added the Excel Interop library to your project! Now, all you have to do is add this using statement to your file:

using Excel = Microsoft.Office.Interop.Excel;
