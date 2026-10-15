# GetDirPro

A folder picker for Visual FoxPro 9, written in VFP: the same call as `GETDIR()`, with a
folder tree that opens at once on any disk, a path the user can type, new folders, and the
theme of your application. One `.prg`.

## Install

```
foxpack add getdirpro
```

```foxpro
SET PROCEDURE TO getdirpro.prg ADDITIVE
```

Without FoxPack, copy `getdirpro.prg` into your project. The tree is the TreeView of the
Windows common controls (`MSCOMCTL.OCX`, installed with VFP 9 and its runtime). Without it,
`GetDir()` falls back to the native `GETDIR()`.

## Example

```foxpro
SET PROCEDURE TO getdirpro.prg ADDITIVE
LOCAL loDir, lcFolder
loDir = CREATEOBJECT("GetDirPro")
lcFolder = loDir.GetDir("C:\Projects", "Where should the project go?")
IF EMPTY(lcFolder)
	? "Cancelled"
ELSE
	? lcFolder    && C:\Projects\Billing\  (with the final backslash)
ENDIF
```

With [FoxStyleSheet](https://github.com/Irwin1985/FoxStyleSheet) loaded, the dialog takes
the theme of your form:

```foxpro
loDir = CREATEOBJECT("GetDirPro")
loDir.cTheme = FssEngineOf(ThisForm).cTheme    && "dark", "light", "Word"...
lcFolder = loDir.GetDir(ThisForm.txtFolder.Value)
```

## API

### GetDirPro.GetDir([cFolder [, cText [, cCaption]]])

Opens the dialog and returns the chosen folder with its final backslash, or `""` when the
user cancels. `cFolder` is where it opens (the current folder of VFP when it is empty or
does not exist), `cText` is a line of text above the tree, and `cCaption` the title of the
window. The returned path keeps the case of each folder on disk.

The user can pick a folder in the tree, type a path, or press **New folder**. A typed path
that does not exist is created, with every missing level, when the user accepts.

### Properties

| Property | Default | What it does |
|---|---|---|
| `cPath` | `""` | What the last `GetDir()` returned |
| `lAllowCreate` | `.T.` | `.F.` lets the user pick only among existing folders and hides **New folder** |
| `cTheme` | `""` | A FoxStyleSheet theme for the dialog: a built-in name, a `.fss` file or the text of a theme. It needs `foxstylesheet.prg` loaded; without it, the dialog keeps its own look |
| `nWarnColor`, `nErrorColor` | orange, red | The colours of the two messages under the path box |
| `cVersion` | `"2.0.0"` | |

### The texts

The dialog is in English. To translate it, change these before `GetDir()`:

| Property | Default |
|---|---|
| `cCapTitle` | `Select folder` |
| `cCapFolder` | `Folder:` |
| `cCapDrive` | `Drive:` |
| `cCapAccept` | `OK` |
| `cCapCancel` | `Cancel` |
| `cCapNew` | `New folder` (the button) |
| `cCapNewName` | `New folder` (the name it proposes) |
| `cCapWillCreate` | `It does not exist yet. It will be created when you accept.` |
| `cCapInvalid` | `That path is not valid.` |
| `cCapCreateFail` | `The folder could not be created. Check the permissions.` |

```foxpro
loDir.cCapTitle  = "Seleccionar carpeta"
loDir.cCapAccept = "Aceptar"
loDir.cCapCancel = "Cancelar"
```

### GetDirPro.FormReady(oForm)

Called with the dialog after it is built and before it is shown. It is empty: bind to it
with `BINDEVENT()`, or subclass `GetDirPro`, to style or change the dialog.

```foxpro
BINDEVENT(loDir, "FormReady", ThisForm, "StyleDialog")
```

### GetDirPro.CreateForm([cFolder [, cText]])

The dialog, built and styled but not shown; `.NULL.` without the common controls.
`GetDir()` calls it and shows the result. Use it to test a screen that opens the picker,
or to show the dialog yourself.

### GetDirPro.EnsureDir(cPath)

Creates a folder and every level it is missing above it (`MD` alone does not). Returns `.T.`
when the folder exists at the end.

### GetDirPro.IsValidPath(cPath)

Whether a typed path can be a folder: a drive (`C:\...`) or a UNC share (`\\server\...`),
and none of the characters Windows refuses. Returns `.T.` or `.F.`.

## Pitfalls

- **The dialog is modal.** `GetDir()` does not return until the user closes it. In tests,
  use `CreateForm()` and drive the dialog without showing it.
- **The tree is an ActiveX control.** A theme does not reach it, so the tree copies the
  colours of the path box after the theme is applied: it follows a dark theme. Its scroll
  bars stay as Windows draws them.
- **An empty folder shows a "+".** The tree reads one level at a time, and a folder shows
  its "+" until it is opened. That is what makes it open at once on a big disk.
- **The title bar of the dialog stays as VFP draws it.** The dialog is a modal window
  inside your form (`ShowWindow = 1`), and Windows does not paint a dark title bar there:
  `darkTitleBar` of a theme only reaches top-level forms.
- **Hidden and system folders are not listed.** A typed path still reaches them.
- **The path ends with a backslash.** `JUSTSTEM()` of it is empty: use
  `JUSTFNAME(LEFT(lcFolder, LEN(lcFolder) - 1))` for the name of the folder.

## Tests

From the repository folder:

```
foxproof run --prg tests\GetDirProTests.prg
```

The styling test applies the `dark` theme when FoxStyleSheet is cloned beside this
repository (`..\FoxStyleSheet`), and skips it otherwise.

## History

- **2.0.0** (2026-10-15). Rewritten. The tree is read one level at a time (1.x walked the
  whole drive when it opened and could hang), folders with a dot in their name are listed,
  hidden and system folders are not, paths keep their case, a path can be typed, folders can
  be created, `SET SAFETY` is left alone, and without the common controls it falls back to
  `GETDIR()`. New: `cTheme`, `FormReady()`, `CreateForm()`, `EnsureDir()`, `IsValidPath()`,
  `lAllowCreate`, the texts as properties. Tests in FoxProof, package for FoxPack, MIT
  license.
- **1.1** (2019-05-10). First public version.

## License

MIT. See [LICENSE](LICENSE).
