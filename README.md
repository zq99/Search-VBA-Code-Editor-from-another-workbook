## Search-VBA-Code-Editor-from-another-workbook

### Purpose

This is demonstration of a class that can search the code in a VBA Editor of project for a particular phrase.

This VBA editor can be in any Microsoft Office Product (Excel/Access/PowerPoint/Word).

The class can be used to form the basis of your own "search VBA code" tool.

The code is useful for finding any hardcoded references when you have several projects to search.

### Requirements

You must have the following VBA Project references installed in the VBE Editor reference window (minimum versions stated):

- Visual Basic For Applications
- Microsoft Excel 14.0 Object Library
- OLE Automation
- Microsoft Scripting Runtime
- Microsoft Visual Basic for Applications Extensibility 5.3
- Microsoft Access 14.0 Object Library
- Microsoft PowerPoint 14.0 Object Library
- Microsoft Word 14.0 Object Library

### Caveats

- You only be able to search the VBA code in applications that do not have the VBA editor locked (unfortunately, there is no way to programmatically unlock the VBA editor)
- For some Excel spreadsheets, you will have to make sure that the option “Trust access to the VBA project object model” has been checked. This can be found under the Macro settings option, within Trust Center.

### Further information

https://datapluscode.com/general/programmatically-search-vba-code/


