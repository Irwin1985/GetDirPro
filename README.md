# GetDirPro ![](images/prg.gif)
Displays the Select Directory dialog box from which you can choose a directory.

### Latest release

**[GetDirPro](https://github.com/Irwin1985/GetDirPro)** - v.1.1 - Release 2019.05.10

Si te gusta mi trabajo puedes apoyarme con un donativo:   
[![DONATE!](http://www.pngall.com/wp-content/uploads/2016/05/PayPal-Donate-Button-PNG-File-180x100.png)](https://www.paypal.com/donate/?hosted_button_id=LXQYXFP77AD2G) 

    Gracias por tu apoyo!


### Usage
```xBase
    *-- Set classLib
    Set Procedure To "GetDirPro.Prg" Additive
    *-- Create object
    loDir = NewObject("GetDirPro", "GetDirPro.prg")
    *-- Prompt dialog box
    ?loDir.getDir([cDirectory [, cText [, cCaption]]])
```

### Parameters 

**cDirectory**
Specifies the directory that is initially displayed in the dialog box. When cDirectory is not specified, the dialog box opens with the Visual FoxPro default directory displayed.

**cText**
Specifies the text for the directory list in the dialog box.

**cCaption**
Specifies the caption to display in the dialog title bar. The Windows default is "Select Directory".

### Return Value
Character
