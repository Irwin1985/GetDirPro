*-- GetDirProTests.prg - GetDirPro 2.0.0
*-- Run from the repository folder: foxproof run --prg tests\GetDirProTests.prg
*--
*-- The dialog is modal, so the tests build it with CreateForm() and drive
*-- it without showing it. Every folder they touch lives in a sandbox
*-- under SYS(2023), deleted afterwards.

#DEFINE GDP_PRG "getdirpro.prg"

* ============================================================ *
* The paths: IsValidPath and EnsureDir
* ============================================================ *
DEFINE CLASS TestPaths AS Custom

	cRoot = ""

	PROCEDURE SetUp
		SET PROCEDURE TO (GDP_PRG) ADDITIVE
		THIS.cRoot = NewSandbox()
	ENDPROC

	PROCEDURE TearDown
		DropSandbox(THIS.cRoot)
	ENDPROC

	PROCEDURE Test_ValidPaths() HELP [Fact, Trait("Category", "Paths")]
		LOCAL loDir
		loDir = CREATEOBJECT("GetDirPro")
		__assert.True(loDir.IsValidPath("C:\Projects\New one"), "a drive path")
		__assert.True(loDir.IsValidPath("c:/projects"), "forward slashes")
		__assert.True(loDir.IsValidPath("\\server\share\folder"), "a UNC path")
	ENDPROC

	PROCEDURE Test_InvalidPaths() HELP [Fact, Trait("Category", "Paths")]
		LOCAL loDir
		loDir = CREATEOBJECT("GetDirPro")
		__assert.False(loDir.IsValidPath("Projects"), "no drive")
		__assert.False(loDir.IsValidPath("C:"), "too short")
		__assert.False(loDir.IsValidPath("C:\a:b"), "a second colon")
		__assert.False(loDir.IsValidPath("C:\what?"), "a question mark")
		__assert.False(loDir.IsValidPath(.NULL.), "not a string")
	ENDPROC

	PROCEDURE Test_EnsureDirCreatesEveryLevel() HELP [Fact, Trait("Category", "Paths")]
		LOCAL loDir, lcPath
		loDir = CREATEOBJECT("GetDirPro")
		lcPath = THIS.cRoot + "one\two\three"
		__assert.True(loDir.EnsureDir(lcPath))
		__assert.True(DIRECTORY(lcPath))
	ENDPROC

	PROCEDURE Test_EnsureDirOnAnExistingFolder() HELP [Fact, Trait("Category", "Paths")]
		LOCAL loDir
		loDir = CREATEOBJECT("GetDirPro")
		__assert.True(loDir.EnsureDir(THIS.cRoot))
		__assert.False(loDir.EnsureDir(""), "an empty path")
	ENDPROC

	PROCEDURE Test_DefaultsAreInEnglish() HELP [Fact, Trait("Category", "Paths")]
		LOCAL loDir
		loDir = CREATEOBJECT("GetDirPro")
		__assert.Equal("Select folder", loDir.cCapTitle)
		__assert.Equal("OK", loDir.cCapAccept)
		__assert.Equal("2.0.0", loDir.cVersion)
	ENDPROC

ENDDEFINE

* ============================================================ *
* The dialog, built and driven without showing it
* ============================================================ *
DEFINE CLASS TestDialog AS Custom

	cRoot = ""

	PROCEDURE SetUp
		SET PROCEDURE TO (GDP_PRG) ADDITIVE
		THIS.cRoot = NewSandbox()
		MD (THIS.cRoot + "Beta")
		MD (THIS.cRoot + "alpha")
		MD (THIS.cRoot + "Zeta")
		MD (THIS.cRoot + "name.with.dots")
		MD (THIS.cRoot + "Beta\Inner")
	ENDPROC

	PROCEDURE TearDown
		DropSandbox(THIS.cRoot)
	ENDPROC

	PROCEDURE Test_OpensOnTheFolderAskedFor() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		loForm = loDir.CreateForm(THIS.cRoot + "Beta\Inner", "Pick one")
		__assert.True(VARTYPE(loForm) == "O", "the common controls are missing")
		__assert.Equal(UPPER(THIS.cRoot + "Beta\Inner\"), UPPER(loForm.txtPath.Value))
		__assert.Equal("Pick one", loForm.lblPrompt.Caption)
		__assert.True(loForm.cmdOk.Enabled, "OK on a folder that exists")
	ENDPROC

	*-- sorted without case, with their case, and the folder with dots there
	PROCEDURE Test_ListsFoldersSortedWithTheirCase() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm, lnN
		LOCAL ARRAY laNames[1]
		loDir = CREATEOBJECT("GetDirPro")
		loForm = loDir.CreateForm(THIS.cRoot, "")
		lnN = loForm.ListFolders(THIS.cRoot, @laNames)
		__assert.Equal(4, lnN)
		__assert.Equal("alpha", laNames[1])
		__assert.Equal("Beta", laNames[2])
		__assert.Equal("name.with.dots", laNames[3])
		__assert.Equal("Zeta", laNames[4])
	ENDPROC

	PROCEDURE Test_ReturnsThePathWithItsCase() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		loForm = loDir.CreateForm(THIS.cRoot + "Beta", "")
		loForm.cmdOk.Click()
		__assert.True(RIGHT(loDir.cPath, 6) == "\Beta\", loDir.cPath)
	ENDPROC

	PROCEDURE Test_TypedFolderIsCreatedOnOk() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm, lcNew
		loDir = CREATEOBJECT("GetDirPro")
		loForm = loDir.CreateForm(THIS.cRoot, "")
		lcNew = THIS.cRoot + "made\here"
		loForm.txtPath.Value = lcNew
		loForm.UpdateState()
		__assert.Equal(loDir.cCapWillCreate, loForm.lblStatus.Caption)
		__assert.Equal(loDir.nWarnColor, loForm.lblStatus.ForeColor)
		loForm.cmdOk.Click()
		__assert.True(DIRECTORY(lcNew), "created with its missing level")
		__assert.Equal(UPPER(ADDBS(lcNew)), UPPER(loDir.cPath))
	ENDPROC

	PROCEDURE Test_InvalidTypedPathDisablesOk() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		loForm = loDir.CreateForm(THIS.cRoot, "")
		loForm.txtPath.Value = "not a path"
		loForm.UpdateState()
		__assert.False(loForm.cmdOk.Enabled)
		__assert.Equal(loDir.cCapInvalid, loForm.lblStatus.Caption)
		__assert.Equal(loDir.nErrorColor, loForm.lblStatus.ForeColor)
	ENDPROC

	PROCEDURE Test_WithoutCreateOnlyExistingFolders() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		loDir.lAllowCreate = .F.
		loForm = loDir.CreateForm(THIS.cRoot, "")
		__assert.False(loForm.cmdNew.Visible, "no New folder button")
		loForm.txtPath.Value = THIS.cRoot + "missing"
		loForm.UpdateState()
		__assert.False(loForm.cmdOk.Enabled)
	ENDPROC

	PROCEDURE Test_NewFolderProposesAFreeName() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		MD (THIS.cRoot + "New folder")
		loForm = loDir.CreateForm(THIS.cRoot, "")
		__assert.Equal("New folder (2)", loForm.ProposeName(THIS.cRoot))
	ENDPROC

	PROCEDURE Test_CancelReturnsNothing() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		loForm = loDir.CreateForm(THIS.cRoot + "Beta", "")
		loForm.cmdCancel.Click()
		__assert.Equal("", loDir.cPath)
	ENDPROC

	PROCEDURE Test_TranslatedTexts() HELP [Fact, Trait("Category", "Dialog")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		loDir.cCapTitle = "Seleccionar carpeta"
		loDir.cCapAccept = "Aceptar"
		loForm = loDir.CreateForm(THIS.cRoot, "")
		__assert.Equal("Seleccionar carpeta", loForm.Caption)
		__assert.Equal("Aceptar", loForm.cmdOk.Caption)
	ENDPROC

ENDDEFINE

* ============================================================ *
* Styling: FormReady() and cTheme
* ============================================================ *
DEFINE CLASS TestStyling AS Custom

	cRoot = ""
	oSeen = .NULL.

	PROCEDURE SetUp
		SET PROCEDURE TO (GDP_PRG) ADDITIVE
		THIS.cRoot = NewSandbox()
	ENDPROC

	PROCEDURE TearDown
		DropSandbox(THIS.cRoot)
		THIS.oSeen = .NULL.
	ENDPROC

	PROCEDURE Seen(toForm)
		THIS.oSeen = toForm
	ENDPROC

	PROCEDURE Test_FormReadyGetsTheDialog() HELP [Fact, Trait("Category", "Styling")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		BINDEVENT(loDir, "FormReady", THIS, "Seen")
		loForm = loDir.CreateForm(THIS.cRoot, "")
		__assert.True(THIS.oSeen == loForm, "FormReady() got the dialog before it was shown")
	ENDPROC

	*-- a theme without foxstylesheet.prg loaded must not stop the dialog
	PROCEDURE Test_ThemeWithoutFoxStyleSheetIsHarmless() HELP [Fact, Trait("Category", "Styling")]
		LOCAL loDir, loForm
		loDir = CREATEOBJECT("GetDirPro")
		loDir.cTheme = "dark"
		loForm = loDir.CreateForm(THIS.cRoot, "")
		__assert.True(VARTYPE(loForm) == "O")
	ENDPROC

	*-- with FoxStyleSheet beside it (as in a FoxPack project), the theme is applied
	PROCEDURE Test_ThemeIsApplied() HELP [Fact, Trait("Category", "Styling")]
		LOCAL loDir, loForm, lcFss
		lcFss = FULLPATH("..\FoxStyleSheet\foxstylesheet.prg")
		IF !FILE(lcFss)
			__assert.True(.T., "FoxStyleSheet is not beside this repository: skipped")
			RETURN
		ENDIF
		SET PROCEDURE TO (lcFss) ADDITIVE
		loDir = CREATEOBJECT("GetDirPro")
		loDir.cTheme = "dark"
		loForm = loDir.CreateForm(THIS.cRoot, "")
		__assert.Equal(RGB(32, 32, 32), loForm.BackColor)
		__assert.Equal(RGB(240, 240, 240), loForm.txtPath.ForeColor)
		*-- the tree, which a theme cannot reach, follows the path box
		__assert.Equal(loForm.txtPath.BackColor, loForm.oleTree.Nodes.Item(1).BackColor)
		__assert.Equal(loForm.txtPath.ForeColor, loForm.oleTree.Nodes.Item(1).ForeColor)
	ENDPROC

ENDDEFINE

* ============================================================ *
* Helpers
* ============================================================ *
FUNCTION NewSandbox
	LOCAL lcRoot
	lcRoot = ADDBS(SYS(2023)) + "gdp_test_" + SYS(2015) + "\"
	MD (lcRoot)
	RETURN lcRoot
ENDFUNC

PROCEDURE DropSandbox
	LPARAMETERS tcRoot
	LOCAL laDirs[1], lnI
	IF EMPTY(tcRoot) OR !DIRECTORY(tcRoot)
		RETURN
	ENDIF
	FOR lnI = 1 TO ADIR(laDirs, tcRoot + "*.*", "D")
		IF "D" $ laDirs[lnI, 5] AND !INLIST(laDirs[lnI, 1], ".", "..")
			DropSandbox(ADDBS(tcRoot + laDirs[lnI, 1]))
		ENDIF
	ENDFOR
	TRY
		RD (tcRoot)
	CATCH
	ENDTRY
ENDPROC
