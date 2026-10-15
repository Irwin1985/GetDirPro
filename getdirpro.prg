*==================================================================
* GetDirPro 2.0.0 -- a folder picker written in Visual FoxPro
*
* Copyright (c) 2019-2026 Irwin Rodríguez. MIT License.
* https://github.com/Irwin1985/GetDirPro
*
* The same call as GETDIR(): GetDir(cFolder, cText, cCaption). Next to
* the 1.x picker, measured on real disks:
*
* 1. THE TREE IS READ ON DEMAND. 1.x walked the whole drive when it
*    opened: on C:\ of a development machine, 75,662 folders and the
*    recursion stuck at 127 levels, the nesting limit of VFP. Now one
*    level is read, and the others when they are opened.
*
* 2. A FOLDER WITH A DOT IN ITS NAME IS THERE. The pattern was "*.",
*    which leaves out "IrwinRodriguez.dev" and the like.
*
* 3. NO HIDDEN OR SYSTEM FOLDERS. That leaves out the compatibility
*    junctions ("Application Data" and company) that point to their own
*    parent, $Recycle.Bin and System Volume Information.
*
* 4. A NODE SHOWS THE NAME, AND THE PATH KEEPS ITS CASE. ADIR() returns
*    names in upper case; the folders come from FileSystemObject, and
*    the path returned is the real one.
*
* 5. A PATH CAN BE TYPED, AND FOLDERS CAN BE CREATED, with every
*    missing level, when the user accepts.
*
* 6. THE ENVIRONMENT IS LEFT ALONE. 1.x turned SET SAFETY off and
*    wrote four BMP files that two pickers at once shared.
*
* 7. WITHOUT THE COMMON CONTROLS IT FALLS BACK TO GETDIR() instead of
*    leaving the user with nothing.
*
* 8. IT TAKES THE THEME OF THE APPLICATION. cTheme links a FoxStyleSheet
*    theme to the dialog, FormReady() lets any code style it, and the
*    tree takes the colours of the path box.
*
* Usage:
*     SET PROCEDURE TO getdirpro.prg ADDITIVE
*     loDir = CREATEOBJECT("GetDirPro")
*     lcFolder = loDir.GetDir("C:\Projects", "Folder of the project")
*==================================================================

#DEFINE GETDIRPRO_VERSION   "2.0.0"
*-- TreeView messages: background, text and line colours
#DEFINE TVM_SETBKCOLOR      4381
#DEFINE TVM_SETTEXTCOLOR    4382
#DEFINE TVM_SETLINECOLOR    4392


DEFINE CLASS GetDirPro AS Custom

    cVersion = GETDIRPRO_VERSION

    *-- What GetDir() returns, with the final backslash; "" when cancelled.
    cPath = ""

    *-- Accept a folder that does not exist yet (creating it), and show
    *-- the New folder button. .F. picks only among the folders there are.
    lAllowCreate = .T.

    *-- A FoxStyleSheet theme for the dialog ("dark", a .fss, the text of
    *-- a theme). Empty: no theme. It needs foxstylesheet.prg loaded.
    cTheme = ""

    *-- The colours of the two messages under the path box.
    nWarnColor  = RGB(196, 110, 0)
    nErrorColor = RGB(200, 30, 30)

    *-- The texts. Change them before calling GetDir() to translate them.
    cCapTitle      = "Select folder"
    cCapFolder     = "Folder:"
    cCapDrive      = "Drive:"
    cCapAccept     = "OK"
    cCapCancel     = "Cancel"
    cCapNew        = "New folder"
    cCapNewName    = "New folder"
    cCapWillCreate = "It does not exist yet. It will be created when you accept."
    cCapInvalid    = "That path is not valid."
    cCapCreateFail = "The folder could not be created. Check the permissions."

    Name = "GetDirPro"


    *--------------------------------------------------------------
    * Opens the dialog and returns the chosen folder, with the final
    * backslash; "" when the user cancels.
    *--------------------------------------------------------------
    FUNCTION GetDir(tcDirectory, tcText, tcCaption)
        LOCAL loForm, lcDir, lcText

        lcDir  = IIF(VARTYPE(tcDirectory) == "C", ALLTRIM(tcDirectory), "")
        lcText = IIF(VARTYPE(tcText) == "C", ALLTRIM(tcText), "")
        IF VARTYPE(tcCaption) == "C" AND !EMPTY(tcCaption)
            THIS.cCapTitle = ALLTRIM(tcCaption)
        ENDIF

        THIS.cPath = ""
        loForm = THIS.CreateForm(lcDir, lcText)

        *-- No tree, no dialog: the native one is better than nothing.
        IF ISNULL(loForm)
            RETURN GETDIR(lcDir, lcText, THIS.cCapTitle)
        ENDIF

        loForm.Show(1)
        RETURN THIS.cPath
    ENDFUNC


    *--------------------------------------------------------------
    * The dialog, built, styled and not shown yet; .NULL. without the
    * common controls. GetDir() shows it; a test can drive it instead.
    *--------------------------------------------------------------
    FUNCTION CreateForm(tcDirectory, tcText)
        LOCAL loForm, loFss

        loForm = .NULL.
        TRY
            loForm = CREATEOBJECT("frmGetDirPro", THIS, ;
                IIF(VARTYPE(tcDirectory) == "C", tcDirectory, ""), ;
                IIF(VARTYPE(tcText) == "C", tcText, ""))
        CATCH
            loForm = .NULL.
        ENDTRY
        IF VARTYPE(loForm) <> "O"
            RETURN .NULL.
        ENDIF

        IF !EMPTY(THIS.cTheme)
            TRY
                loFss = CREATEOBJECT("FoxStyleSheet")
                loFss.link(THIS.cTheme, loForm)
            CATCH
                *-- without foxstylesheet.prg the dialog keeps its own look
            ENDTRY
        ENDIF
        THIS.FormReady(loForm)
        loForm.MatchTreeColors()
        RETURN loForm
    ENDFUNC


    *--------------------------------------------------------------
    * Called with the dialog before it is shown. Empty: bind to it
    * (BINDEVENT) or subclass to style or change the dialog.
    *--------------------------------------------------------------
    PROCEDURE FormReady(toForm)
    ENDPROC


    *--------------------------------------------------------------
    * Creates a folder and every level it is missing above it (MD alone
    * does not). .T. when the folder exists at the end. Public because
    * any screen that lets the user type a path needs the same thing.
    *--------------------------------------------------------------
    FUNCTION EnsureDir(tcPath)
        LOCAL lcPath, lcParent

        lcPath = ALLTRIM(IIF(VARTYPE(tcPath) == "C", tcPath, ""))
        IF EMPTY(lcPath)
            RETURN .F.
        ENDIF

        lcPath = ADDBS(lcPath)
        IF DIRECTORY(lcPath)
            RETURN .T.
        ENDIF

        lcParent = ADDBS(JUSTPATH(LEFT(lcPath, LEN(lcPath) - 1)))
        IF LEN(lcParent) < LEN(lcPath) AND LEN(lcParent) > 2
            IF !THIS.EnsureDir(lcParent)
                RETURN .F.
            ENDIF
        ENDIF

        TRY
            MD (lcPath)
        CATCH
        ENDTRY

        RETURN DIRECTORY(lcPath)
    ENDFUNC


    *--------------------------------------------------------------
    * Whether a typed path can be a folder: a drive or a UNC share, and
    * none of the characters Windows refuses. A colon is legal after
    * the drive letter and nowhere else.
    *--------------------------------------------------------------
    FUNCTION IsValidPath(tcPath)
        LOCAL lcP, i, lcBad

        lcP = ALLTRIM(IIF(VARTYPE(tcPath) == "C", tcPath, ""))
        IF LEN(lcP) < 3
            RETURN .F.
        ENDIF

        IF !((ISALPHA(LEFT(lcP, 1)) AND SUBSTR(lcP, 2, 1) == ":" ;
              AND SUBSTR(lcP, 3, 1) $ "\/") OR LEFT(lcP, 2) == "\\")
            RETURN .F.
        ENDIF

        lcBad = '*?"<>|'
        FOR i = 1 TO LEN(lcBad)
            IF SUBSTR(lcBad, i, 1) $ lcP
                RETURN .F.
            ENDIF
        ENDFOR

        IF ":" $ SUBSTR(lcP, 3)
            RETURN .F.
        ENDIF

        RETURN .T.
    ENDFUNC

ENDDEFINE


*==================================================================
* The dialog. Reached through GetDirPro.GetDir() or CreateForm().
*==================================================================
DEFINE CLASS frmGetDirPro AS Form

    DataSession = 2
    Caption     = "Select folder"
    Width       = 560
    Height      = 460
    MinWidth    = 460
    MinHeight   = 380
    AutoCenter  = .T.
    BorderStyle = 3           && resizable: a folder tree likes room
    MaxButton   = .F.
    MinButton   = .F.
    WindowType  = 1           && modal
    ShowWindow  = 1
    Name        = "frmGetDirPro"

    oOwner  = .NULL.          && the GetDirPro that opened it
    oPaths  = .NULL.          && node key -> real path
    oKeys   = .NULL.          && path in upper case -> node key
    oLoaded = .NULL.          && keys whose children were read
    nSeq    = 0               && key counter
    cSelKey = ""              && node drawn as selected
    lReady  = .F.             && .F. while it is built: events wait
    nTreeBack = -1            && colours of the nodes, -1 = Windows'
    nTreeFore = -1


    ADD OBJECT lblPrompt AS label WITH ;
        Caption = "", BackStyle = 0, ;
        Left = 12, Top = 12, Width = 360, Height = 18, ;
        Name = "lblPrompt"

    ADD OBJECT lblDrive AS label WITH ;
        Caption = "", BackStyle = 0, Alignment = 1, ;
        Left = 380, Top = 14, Width = 60, Height = 16, ;
        Name = "lblDrive"

    ADD OBJECT cmbDrive AS combobox WITH ;
        Style = 2, RowSourceType = 0, ;
        Left = 444, Top = 10, Width = 104, Height = 24, ;
        Name = "cmbDrive"

    ADD OBJECT oleTree AS olecontrol WITH ;
        oleclass = "MSComctlLib.TreeCtrl.2", ;
        Left = 12, Top = 42, Width = 536, Height = 290, ;
        Name = "oleTree"

    ADD OBJECT oleImages AS olecontrol WITH ;
        oleclass = "MSComctlLib.ImageListCtrl.2", ;
        Left = -100, Top = -100, Width = 32, Height = 32, ;
        Visible = .F., ;
        Name = "oleImages"

    ADD OBJECT lblPath AS label WITH ;
        Caption = "", BackStyle = 0, ;
        Left = 12, Top = 342, Width = 200, Height = 16, ;
        Name = "lblPath"

    ADD OBJECT txtPath AS textbox WITH ;
        Left = 12, Top = 360, Width = 536, Height = 24, ;
        Value = "", ;
        Name = "txtPath"

    ADD OBJECT lblStatus AS label WITH ;
        Caption = "", BackStyle = 0, ;
        Left = 12, Top = 390, Width = 536, Height = 16, ;
        Name = "lblStatus"

    ADD OBJECT cmdNew AS commandbutton WITH ;
        Caption = "", Enabled = .F., ;
        Left = 12, Top = 412, Width = 130, Height = 28, ;
        Name = "cmdNew"

    ADD OBJECT cmdOk AS commandbutton WITH ;
        Caption = "", Default = .T., Enabled = .F., ;
        Left = 356, Top = 412, Width = 92, Height = 28, ;
        Name = "cmdOk"

    ADD OBJECT cmdCancel AS commandbutton WITH ;
        Caption = "", Cancel = .T., ;
        Left = 456, Top = 412, Width = 92, Height = 28, ;
        Name = "cmdCancel"


    *--------------------------------------------------------------
    PROCEDURE Init(toOwner, tcDirectory, tcText)
        LOCAL lcStart

        IF VARTYPE(toOwner) <> "O"
            RETURN .F.
        ENDIF
        THIS.oOwner = toOwner

        *-- Without the common controls there is nothing to do here. An
        *-- OleControl that could not be created is not an object.
        IF TYPE("THIS.oleTree") <> "O" OR ISNULL(THIS.oleTree)
            RETURN .F.
        ENDIF
        IF TYPE("THIS.oleImages") <> "O" OR ISNULL(THIS.oleImages)
            RETURN .F.
        ENDIF

        THIS.oPaths  = CREATEOBJECT("Collection")
        THIS.oKeys   = CREATEOBJECT("Collection")
        THIS.oLoaded = CREATEOBJECT("Collection")

        THIS.Caption            = toOwner.cCapTitle
        THIS.lblPath.Caption    = toOwner.cCapFolder
        THIS.lblDrive.Caption   = toOwner.cCapDrive
        THIS.cmdOk.Caption      = toOwner.cCapAccept
        THIS.cmdCancel.Caption  = toOwner.cCapCancel
        THIS.cmdNew.Caption     = toOwner.cCapNew

        THIS.lblPrompt.Caption = IIF(VARTYPE(tcText) == "C", ALLTRIM(tcText), "")
        THIS.lblPrompt.Visible = !EMPTY(THIS.lblPrompt.Caption)

        IF !toOwner.lAllowCreate
            THIS.cmdNew.Visible = .F.
        ENDIF

        *-- Without icons it goes on: a tree without pictures still works.
        THIS.LoadIcons()

        THIS.FillDrives()

        *-- Where to start: the folder asked for if it exists, and else
        *-- the current folder of VFP.
        lcStart = ""
        IF VARTYPE(tcDirectory) == "C" AND !EMPTY(tcDirectory)
            IF DIRECTORY(ALLTRIM(tcDirectory))
                lcStart = ADDBS(ALLTRIM(tcDirectory))
            ENDIF
        ENDIF
        IF EMPTY(lcStart)
            lcStart = ADDBS(SYS(5) + CURDIR())
        ENDIF

        THIS.Relayout()
        THIS.lReady = .T.

        THIS.SelectDrive(LEFT(lcStart, 2))
        THIS.NavigateTo(lcStart)
        THIS.UpdateState()

        RETURN .T.
    ENDPROC


    *--------------------------------------------------------------
    * The two folder icons travel here in base64, so that the library
    * is one .prg. Each picture is written to a file of its own and
    * deleted as soon as LoadPicture has it: two dialogs at once do not
    * share files, and nothing is left behind if the process dies.
    *--------------------------------------------------------------
    FUNCTION LoadIcons()
        LOCAL lcOpen, lcClose, llOk

        llOk = .F.
        TEXT TO m.lcOpen NOSHOW
Qk02AwAAAAAAADYAAAAoAAAAEAAAABAAAAABABgAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAA////////////////////////////////////////////////////////////////v+75quj3quj3quj3quj3quj3quj3quj3quj3quj3quj3quj3ren49v3+////////W9LxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vsur4////////XNDxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vb9jy////////QcfuTM7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7v4/f8////QMbukt/zSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7voOX3////QMbu0O72Ss7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vX9Tx/v//QMbu3e/0V9LxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7v0fP7QMbu3O/0p+T0Ss7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vjuD1QMbu3O/02fD1TM7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vYtXxKr/su9jfzuPo2evw0OTp1unu0uXq1Ojt1Ojt0uXq1unu0OTp2ezx3fD15fT4+f3+FrrrmdHiyN/k3O/03O/03O/03O/03O/03O/03O/03O/03O/03O/03O/06fX4////Frrrks3fhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/R9/v8////FrrrKb7qNMDrNMDrNMDrN8Dqzu739Pn79Pn79Pn79Pn79Pn79Pn79Pn7////////rOb4kd/2kd/2kd/2kd/2s+n4////////////////////////////////////////////////////////////////////////////////////////////////////////
        ENDTEXT

        TEXT TO m.lcClose NOSHOW
Qk02AwAAAAAAADYAAAAoAAAAEAAAABAAAAABABgAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAA////////////////////////////////////////////////////////////////v+75quj3quj3quj3quj3quj3quj3quj3quj3quj3quj3quj3ren49v3+////////WtHwSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vsur4////////W9DxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vb9jy////////LcHtS87vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7v4/f8////FrrrXtLxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7voOX3////FrrrU83wSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vX9Tx/v//FrrrIL3sT9DwSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7v0fP7FrrrFrrrYdLySs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vjuD1FrrrFrrrRcjvS87vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vYtXxFrrrFrrrFrrrIL3rI77sI77sI77sI77sI77sI77sI77sI77sI77sI77srOf4+f3+FrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrr+f4////FrrrFrrrFrrrFrrrFrrrFrrrMMLte9j0e9j0e9j0e9j0e9j0e9j0ftn07Pn9////FrrrFrrrFrrrFrrrFrrrGbvrx+76////////////////////////////////////rOb4kd/2kd/2kd/2kd/2suj4////////////////////////////////////////////////////////////////////////////////////////////////////////
        ENDTEXT

        TRY
            THIS.oleImages.ImageHeight = 16
            THIS.oleImages.ImageWidth  = 16
            *-- 1 = open (the selected folder), 2 = closed (the others)
            THIS.oleImages.ListImages.Add(, "open",  THIS.PictureFrom(m.lcOpen))
            THIS.oleImages.ListImages.Add(, "close", THIS.PictureFrom(m.lcClose))
            THIS.oleTree.ImageList = THIS.oleImages.Object
            llOk = .T.
        CATCH
            llOk = .F.
        ENDTRY

        RETURN llOk
    ENDFUNC


    FUNCTION PictureFrom(tcBase64)
        LOCAL lcFile, loPic

        lcFile = ADDBS(SYS(2023)) + "gdp" + SYS(2015) + ".bmp"
        STRTOFILE(STRCONV(tcBase64, 14), lcFile)
        loPic = LOADPICTURE(lcFile)

        TRY
            DELETE FILE (lcFile)
        CATCH
        ENDTRY

        RETURN loPic
    ENDFUNC


    *--------------------------------------------------------------
    * The tree is an ActiveX control: a theme does not reach it. It
    * takes the colours of the path box, so that it follows the theme
    * of the dialog, the dark one included.
    *--------------------------------------------------------------
    PROCEDURE MatchTreeColors
        LOCAL lnHwnd, loNode
        THIS.nTreeBack = THIS.txtPath.BackColor
        THIS.nTreeFore = THIS.txtPath.ForeColor
        TRY
            DECLARE INTEGER SendMessage IN user32 AS GdpSendMessage ;
                INTEGER hWnd, INTEGER Msg, INTEGER wParam, INTEGER lParam
            lnHwnd = THIS.oleTree.hWnd
            =GdpSendMessage(lnHwnd, TVM_SETBKCOLOR, 0, THIS.nTreeBack)
            =GdpSendMessage(lnHwnd, TVM_SETTEXTCOLOR, 0, THIS.nTreeFore)
            =GdpSendMessage(lnHwnd, TVM_SETLINECOLOR, 0, THIS.nTreeFore)
        CATCH
            *-- the tree keeps the colours of Windows
        ENDTRY
        *-- each node carries colours of its own, white and black by default
        TRY
            FOR EACH loNode IN THIS.oleTree.Nodes
                THIS.PaintNode(loNode)
            ENDFOR
        CATCH
        ENDTRY
    ENDPROC


    PROCEDURE PaintNode(toNode)
        IF THIS.nTreeBack < 0 OR VARTYPE(toNode) <> "O"
            RETURN
        ENDIF
        TRY
            toNode.BackColor = THIS.nTreeBack
            toNode.ForeColor = THIS.nTreeFore
        CATCH
        ENDTRY
    ENDPROC


    *--------------------------------------------------------------
    * Drives: the ones that answer. An empty drive or a mapped drive
    * that is down does not hang the dialog: DIRECTORY() does not block.
    *--------------------------------------------------------------
    PROCEDURE FillDrives
        LOCAL i, lcVol
        FOR i = 65 TO 90
            lcVol = CHR(i) + ":\"
            IF DIRECTORY(lcVol)
                THIS.cmbDrive.AddItem(CHR(i) + ":")
            ENDIF
        ENDFOR
    ENDPROC


    PROCEDURE SelectDrive(tcDrive)
        LOCAL i, lcD
        lcD = UPPER(LEFT(ALLTRIM(tcDrive) + "  ", 2))
        FOR i = 1 TO THIS.cmbDrive.ListCount
            IF UPPER(THIS.cmbDrive.List(i)) == lcD
                THIS.cmbDrive.ListIndex = i
                THIS.SetRoot(lcD)
                RETURN
            ENDIF
        ENDFOR
        IF THIS.cmbDrive.ListCount > 0
            THIS.cmbDrive.ListIndex = 1
            THIS.SetRoot(THIS.cmbDrive.List(1))
        ENDIF
    ENDPROC


    *--------------------------------------------------------------
    * Plants the root of a drive and reads ONLY its first level.
    *--------------------------------------------------------------
    PROCEDURE SetRoot(tcDrive)
        LOCAL lcRoot, lcKey, loNode

        lcRoot = ADDBS(UPPER(ALLTRIM(tcDrive)))

        THIS.oleTree.Nodes.Clear
        THIS.oPaths  = CREATEOBJECT("Collection")
        THIS.oKeys   = CREATEOBJECT("Collection")
        THIS.oLoaded = CREATEOBJECT("Collection")
        THIS.cSelKey = ""

        lcKey = THIS.NewKey()
        THIS.Remember(lcKey, lcRoot)

        loNode = THIS.oleTree.Nodes.Add(, 1, lcKey, lcRoot)
        THIS.SetImage(loNode, 2)
        THIS.PaintNode(loNode)

        THIS.LoadChildren(loNode)
        loNode.Expanded = .T.
    ENDPROC


    *--------------------------------------------------------------
    * The heart of it: ONE level is read, and only once.
    *
    * Each child is born with a placeholder child, so that the tree
    * draws the "+" without looking inside. When it is opened, the
    * placeholder goes and the real children are read. Checking whether
    * each sibling has children would cost one listing per sibling; the
    * price of not doing it is that an empty folder shows a "+" until it
    * is opened, as the Windows Explorer did for years.
    *--------------------------------------------------------------
    PROCEDURE LoadChildren(toNode)
        LOCAL lcKey, lcPath, lcDummy, i, lnTot, lcName, lcChildKey, loChild
        LOCAL ARRAY laNames[1]

        lcKey = toNode.Key
        IF THIS.oLoaded.GetKey(lcKey) > 0
            RETURN
        ENDIF
        THIS.oLoaded.Add(.T., lcKey)

        lcPath = THIS.PathOf(lcKey)
        IF EMPTY(lcPath)
            RETURN
        ENDIF

        lcDummy = lcKey + "@d"
        TRY
            THIS.oleTree.Nodes.Remove(lcDummy)
        CATCH
        ENDTRY

        lnTot = THIS.ListFolders(lcPath, @laNames)

        FOR i = 1 TO lnTot
            lcName = laNames[i]

            lcChildKey = THIS.NewKey()
            THIS.Remember(lcChildKey, ADDBS(ADDBS(lcPath) + lcName))

            loChild = THIS.oleTree.Nodes.Add(lcKey, 4, lcChildKey, lcName)
            THIS.SetImage(loChild, 2)
            THIS.PaintNode(loChild)
            THIS.AddDummy(lcChildKey)
        ENDFOR
    ENDPROC


    *--------------------------------------------------------------
    * The names of the subfolders, sorted and WITH THEIR CASE.
    *
    * They come from FileSystemObject and not from ADIR(), because
    * ADIR() returns names in upper case: the tree would show FOXAGENT
    * where the folder is FoxAgent, and that path is what the caller
    * would store. Hidden and system folders (attributes 2 and 4) are
    * left out. Without FileSystemObject, ADIR(): a tree in capitals is
    * better than no tree.
    *--------------------------------------------------------------
    FUNCTION ListFolders(tcPath, taNames)
        LOCAL loFSO, loFolder, loSub, lnCount, i, lnTot, llFso
        LOCAL ARRAY laD[1, 1]

        lnCount = 0
        llFso   = .F.

        TRY
            loFSO    = CREATEOBJECT("Scripting.FileSystemObject")
            loFolder = loFSO.GetFolder(ADDBS(tcPath))
            llFso    = .T.
        CATCH
            llFso = .F.
        ENDTRY

        IF llFso
            TRY
                FOR EACH loSub IN loFolder.SubFolders
                    IF BITAND(loSub.Attributes, 6) > 0
                        LOOP
                    ENDIF
                    lnCount = lnCount + 1
                    DIMENSION taNames[lnCount]
                    taNames[lnCount] = loSub.Name
                ENDFOR
            CATCH
                *-- A folder that cannot be entered stops the listing:
                *-- what was read before stays, which beats nothing.
            ENDTRY
        ELSE
            TRY
                lnTot = ADIR(laD, ADDBS(tcPath) + "*.*", "D")
            CATCH
                lnTot = 0
            ENDTRY
            FOR i = 1 TO lnTot
                IF ATC("D", laD[i, 5]) = 0
                    LOOP
                ENDIF
                IF laD[i, 1] == "." OR laD[i, 1] == ".."
                    LOOP
                ENDIF
                lnCount = lnCount + 1
                DIMENSION taNames[lnCount]
                taNames[lnCount] = laD[i, 1]
            ENDFOR
        ENDIF

        *-- The last 1 ignores case: otherwise Zeta goes before alpha.
        IF lnCount > 1
            ASORT(taNames, 1, lnCount, 0, 1)
        ENDIF

        RETURN lnCount
    ENDFUNC


    PROCEDURE AddDummy(tcKey)
        TRY
            THIS.PaintNode(THIS.oleTree.Nodes.Add(tcKey, 4, tcKey + "@d", "..."))
        CATCH
        ENDTRY
    ENDPROC


    *--------------------------------------------------------------
    * Opens the tree down to a path and leaves it selected.
    *--------------------------------------------------------------
    FUNCTION NavigateTo(tcPath)
        LOCAL lcPath, lcAcum, lnI, lnN, lcKey, loNode
        LOCAL ARRAY laParts[1]

        lcPath = ADDBS(ALLTRIM(tcPath))
        IF !DIRECTORY(lcPath)
            RETURN .F.
        ENDIF

        lnN = ALINES(laParts, STRTRAN(lcPath, "\", CHR(13)))
        IF lnN < 1
            RETURN .F.
        ENDIF

        lcAcum = ADDBS(laParts[1])
        lcKey  = THIS.KeyOf(lcAcum)
        IF EMPTY(lcKey)
            RETURN .F.
        ENDIF

        loNode = THIS.oleTree.Nodes.Item(lcKey)
        THIS.LoadChildren(loNode)
        loNode.Expanded = .T.

        FOR lnI = 2 TO lnN
            IF EMPTY(ALLTRIM(laParts[lnI]))
                LOOP
            ENDIF
            lcAcum = ADDBS(lcAcum + ALLTRIM(laParts[lnI]))
            lcKey  = THIS.KeyOf(lcAcum)
            IF EMPTY(lcKey)
                EXIT
            ENDIF
            loNode = THIS.oleTree.Nodes.Item(lcKey)
            THIS.LoadChildren(loNode)
            loNode.Expanded = .T.
        ENDFOR

        THIS.Highlight(loNode)
        loNode.Selected = .T.
        TRY
            loNode.EnsureVisible()
        CATCH
        ENDTRY

        THIS.txtPath.Value = THIS.PathOf(loNode.Key)

        RETURN .T.
    ENDFUNC


    *--------------------------------------------------------------
    * Keys and paths. A node shows the NAME of the folder; the real
    * path, with its case, lives here, and it is the one returned.
    *--------------------------------------------------------------
    FUNCTION NewKey()
        THIS.nSeq = THIS.nSeq + 1
        RETURN "n" + TRANSFORM(THIS.nSeq)
    ENDFUNC


    PROCEDURE Remember(tcKey, tcPath)
        LOCAL lnIdx

        THIS.oPaths.Add(tcPath, tcKey)

        *-- A path comes again when a node is reloaded: the collection
        *-- refuses a duplicate key, so the old one goes.
        lnIdx = THIS.oKeys.GetKey(UPPER(tcPath))
        IF lnIdx > 0
            THIS.oKeys.Remove(lnIdx)
        ENDIF
        THIS.oKeys.Add(tcKey, UPPER(tcPath))
    ENDPROC


    FUNCTION PathOf(tcKey)
        IF THIS.oPaths.GetKey(tcKey) > 0
            RETURN THIS.oPaths.Item(tcKey)
        ENDIF
        RETURN ""
    ENDFUNC


    FUNCTION KeyOf(tcPath)
        IF THIS.oKeys.GetKey(UPPER(ADDBS(ALLTRIM(tcPath)))) > 0
            RETURN THIS.oKeys.Item(UPPER(ADDBS(ALLTRIM(tcPath))))
        ENDIF
        RETURN ""
    ENDFUNC


    PROCEDURE SetImage(toNode, tnImage)
        TRY
            toNode.Image = tnImage
        CATCH
        ENDTRY
    ENDPROC


    *-- Only TWO nodes are repainted, not the whole tree.
    PROCEDURE Highlight(toNode)
        IF !EMPTY(THIS.cSelKey)
            TRY
                THIS.SetImage(THIS.oleTree.Nodes.Item(THIS.cSelKey), 2)
            CATCH
            ENDTRY
        ENDIF
        THIS.SetImage(toNode, 1)
        THIS.cSelKey = toNode.Key
    ENDPROC


    *--------------------------------------------------------------
    * The buttons, and the "does not exist" message.
    *--------------------------------------------------------------
    PROCEDURE UpdateState
        LOCAL lcPath, llExists

        lcPath = ALLTRIM(THIS.txtPath.Value)
        THIS.lblStatus.Caption = ""

        IF EMPTY(lcPath)
            THIS.cmdOk.Enabled  = .F.
            THIS.cmdNew.Enabled = .F.
            RETURN
        ENDIF

        llExists = DIRECTORY(lcPath)

        *-- A new folder hangs from one that exists.
        THIS.cmdNew.Enabled = llExists AND THIS.oOwner.lAllowCreate

        IF llExists
            THIS.cmdOk.Enabled = .T.
            RETURN
        ENDIF

        IF THIS.oOwner.lAllowCreate AND THIS.oOwner.IsValidPath(lcPath)
            THIS.lblStatus.Caption   = THIS.oOwner.cCapWillCreate
            THIS.lblStatus.ForeColor = THIS.oOwner.nWarnColor
            THIS.cmdOk.Enabled       = .T.
        ELSE
            THIS.lblStatus.Caption   = THIS.oOwner.cCapInvalid
            THIS.lblStatus.ForeColor = THIS.oOwner.nErrorColor
            THIS.cmdOk.Enabled       = .F.
        ENDIF
    ENDPROC


    *-- Reads the children of a folder again, without reloading the tree.
    PROCEDURE ReloadNode(tcPath)
        LOCAL lcKey, loNode, lnIdx

        lcKey = THIS.KeyOf(tcPath)
        IF EMPTY(lcKey)
            RETURN
        ENDIF

        TRY
            loNode = THIS.oleTree.Nodes.Item(lcKey)
        CATCH
            RETURN
        ENDTRY

        DO WHILE loNode.Children > 0
            TRY
                THIS.oleTree.Nodes.Remove(loNode.Child.Index)
            CATCH
                EXIT
            ENDTRY
        ENDDO

        lnIdx = THIS.oLoaded.GetKey(lcKey)
        IF lnIdx > 0
            THIS.oLoaded.Remove(lnIdx)
        ENDIF

        THIS.LoadChildren(loNode)
        loNode.Expanded = .T.
    ENDPROC


    *--------------------------------------------------------------
    * Events
    *--------------------------------------------------------------
    PROCEDURE oleTree.Expand
        LPARAMETERS toNode
        IF !THISFORM.lReady
            RETURN
        ENDIF
        THISFORM.LoadChildren(toNode)
    ENDPROC


    PROCEDURE oleTree.NodeClick
        LPARAMETERS toNode
        IF !THISFORM.lReady
            RETURN
        ENDIF
        THISFORM.Highlight(toNode)
        THISFORM.txtPath.Value = THISFORM.PathOf(toNode.Key)
        THISFORM.UpdateState()
    ENDPROC


    PROCEDURE oleTree.DblClick
        IF !THISFORM.lReady
            RETURN
        ENDIF
        *-- A double click on a folder that exists accepts it.
        IF THISFORM.cmdOk.Enabled AND DIRECTORY(ALLTRIM(THISFORM.txtPath.Value))
            THISFORM.cmdOk.Click()
        ENDIF
    ENDPROC


    PROCEDURE cmbDrive.InteractiveChange
        IF !THISFORM.lReady
            RETURN
        ENDIF
        THISFORM.SetRoot(THIS.Value)
        THISFORM.txtPath.Value = ADDBS(UPPER(ALLTRIM(THIS.Value)))
        THISFORM.UpdateState()
    ENDPROC


    *-- The path box does NOT move the tree, on purpose: OK is the
    *-- default button, so Enter accepts the typed path, and moving the
    *-- tree on LostFocus means rebuilding the whole ActiveX.
    PROCEDURE txtPath.InteractiveChange
        IF !THISFORM.lReady
            RETURN
        ENDIF
        THISFORM.UpdateState()
    ENDPROC


    *-- No InputBox, on purpose: a dialog inside a dialog cannot be
    *-- tested without a human. It proposes the new path in the box with
    *-- the name SELECTED to type over it; OK creates it, the same way as
    *-- a typed path that does not exist.
    PROCEDURE cmdNew.Click
        LOCAL lcBase, lcProp

        lcBase = ALLTRIM(THISFORM.txtPath.Value)
        IF !DIRECTORY(lcBase)
            RETURN
        ENDIF

        lcProp = THISFORM.ProposeName(lcBase)
        THISFORM.txtPath.Value = ADDBS(lcBase) + lcProp

        TRY
            THISFORM.txtPath.SetFocus()
            THISFORM.txtPath.SelStart  = LEN(ADDBS(lcBase))
            THISFORM.txtPath.SelLength = LEN(lcProp)
        CATCH
        ENDTRY

        THISFORM.UpdateState()
    ENDPROC


    *-- "New folder", "New folder (2)"... until one is free.
    FUNCTION ProposeName(tcBase)
        LOCAL lcName, lnI

        lcName = THIS.oOwner.cCapNewName
        IF !DIRECTORY(ADDBS(tcBase) + lcName)
            RETURN lcName
        ENDIF

        FOR lnI = 2 TO 99
            IF !DIRECTORY(ADDBS(tcBase) + lcName + " (" + TRANSFORM(lnI) + ")")
                RETURN lcName + " (" + TRANSFORM(lnI) + ")"
            ENDIF
        ENDFOR

        RETURN lcName
    ENDFUNC


    PROCEDURE cmdOk.Click
        LOCAL lcPath

        lcPath = ALLTRIM(THISFORM.txtPath.Value)
        IF EMPTY(lcPath)
            RETURN
        ENDIF

        IF !DIRECTORY(lcPath)
            IF !THISFORM.oOwner.lAllowCreate
                RETURN
            ENDIF
            IF !THISFORM.oOwner.EnsureDir(lcPath)
                THISFORM.lblStatus.Caption   = THISFORM.oOwner.cCapCreateFail
                THISFORM.lblStatus.ForeColor = THISFORM.oOwner.nErrorColor
                RETURN
            ENDIF
        ENDIF

        THISFORM.oOwner.cPath = ADDBS(lcPath)
        THISFORM.Release()
    ENDPROC


    PROCEDURE cmdCancel.Click
        THISFORM.oOwner.cPath = ""
        THISFORM.Release()
    ENDPROC


    *--------------------------------------------------------------
    * Layout, from the current size, so that the window can stretch.
    *--------------------------------------------------------------
    PROCEDURE Relayout
        LOCAL lnW, lnH, lnBase

        lnW = THIS.Width
        lnH = THIS.Height

        THIS.lblPrompt.Left  = 12
        THIS.lblPrompt.Top   = 12
        THIS.lblPrompt.Width = MAX(lnW - 190, 80)

        THIS.lblDrive.Left = lnW - 180
        THIS.lblDrive.Top  = 14
        THIS.cmbDrive.Left = lnW - 116
        THIS.cmbDrive.Top  = 10

        THIS.oleTree.Left   = 12
        THIS.oleTree.Top    = 42
        THIS.oleTree.Width  = MAX(lnW - 24, 200)
        THIS.oleTree.Height = MAX(lnH - 170, 80)

        lnBase = THIS.oleTree.Top + THIS.oleTree.Height + 10

        THIS.lblPath.Left  = 12
        THIS.lblPath.Top   = lnBase

        THIS.txtPath.Left  = 12
        THIS.txtPath.Top   = lnBase + 18
        THIS.txtPath.Width = MAX(lnW - 24, 200)

        THIS.lblStatus.Left  = 12
        THIS.lblStatus.Top   = lnBase + 48
        THIS.lblStatus.Width = MAX(lnW - 24, 200)

        THIS.cmdNew.Left    = 12
        THIS.cmdNew.Top     = lnBase + 70
        THIS.cmdOk.Left     = lnW - 204
        THIS.cmdOk.Top      = lnBase + 70
        THIS.cmdCancel.Left = lnW - 104
        THIS.cmdCancel.Top  = lnBase + 70
    ENDPROC


    PROCEDURE Resize
        IF !THIS.lReady
            RETURN
        ENDIF
        THIS.Relayout()
    ENDPROC

ENDDEFINE
