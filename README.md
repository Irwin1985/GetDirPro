# GetDirPro ![](images/prg.gif)
Displays the Select Directory dialog box from which you can choose a directory.

### Latest release

**[GetDirPro](https://github.com/Irwin1985/GetDirPro)** - v.1.1 - Release 2019.05.10

### Usage
<pre>*-- Create object
    loDir = CreateObject("GetDirPro", "GetDirPro.prg")
    *-- Prompt dialog box
    ?loDir.getDir([cDirectory [, cText [, cCaption]]])
</pre>

### Parameters 

**cDirectory**
Specifies the directory that is initially displayed in the dialog box. When cDirectory is not specified, the dialog box opens with the Visual FoxPro default directory displayed.

**cText**
Specifies the text for the directory list in the dialog box.

**cCaption**
Specifies the caption to display in the dialog title bar. The Windows default is "Select Directory".

### Return Value
Character
