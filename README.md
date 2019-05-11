# GetDirPro ![](images/fbc_icon.png)
Displays the Select Directory dialog box from which you can choose a directory.
Usage:

    loDir = CreateObject("GetDirPro", "GetDirPro.prg")
    ?loDir.getDir([cDirectory [, cText [, cCaption]]])

**PARAMETERS** 

    Same as GETDIR() native function.

    [tcDirectory] = Default directory.
    [tcText]      = Alternative text (above treeview)
    [tcCaption]   = Windows title.
