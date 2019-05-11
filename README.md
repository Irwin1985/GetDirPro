# GetDirPro ![](images/prg.gif)
Displays the Select Directory dialog box from which you can choose a directory.

### Latest release

**[GetDirPro_v_1_1](/GetDirPro_v_1_1/)** - v.1.1 - Release 2019.05.10

Usage:

    loDir = CreateObject("GetDirPro", "GetDirPro.prg")
    ?loDir.getDir([cDirectory [, cText [, cCaption]]])

**PARAMETERS** 

    Same as GETDIR() native function.

    [tcDirectory] = Default directory.
    [tcText]      = Alternative text (above treeview)
    [tcCaption]   = Windows title.
