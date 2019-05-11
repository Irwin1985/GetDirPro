*---------------------------------------------------------------------------------------------------------------*
*
* @title:		Librería GetDirPro
* @description:	Selector de Directorios hecho 100% en Visual FoxPro 9.0
*
* @version:		1.2
* @author:		Irwin Rodríguez
* @email:		rodriguez.irwin@gmail.com
* @license:		MIT
*
* @usage:
* loDir = CreateObject("GetDirPro", "GetDirPro.prg")
* ?loDir.GetDir()
*
* -------------------------------------------------------------------------
* Version Log:
* Release 2019-05-10	v.1.2		- Liberación formal en https://github.com/Irwin1985/GetDirPro
*---------------------------------------------------------------------------------------------------------------*

Define Class GetDirPro As Custom
	Procedure Getdir
		Lparameters tcDirectory As String, tcText As String, tcCaption As String
		loForm = Createobject("frmGetDir", tcDirectory, tcText, tcCaption)
		loForm.Show(1)
		cPath = loForm.cPath
		Release loForm
		Return cPath
	Endproc
Enddefine

*-- Internal Usage
Define Class frmGetDir As Form

	DataSession = 2
	BorderStyle = 2
	Height 		= 463
	Width 		= 613
	DoCreate 	= .T.
	AutoCenter 	= .T.
	Caption 	= "Form"
	MaxButton 	= .F.
	Closable	= .F.
	MinButton 	= .F.
	BackColor 	= Rgb(250,250,250)
	cPath 		= ""

	Add Object oletreeview As OleControl With ;
		oleclass	= "MSComctlLib.TreeCtrl.2", ;
		Top 		= 40, ;
		Left 		= 12, ;
		Height 		= 360, ;
		Width 		= 590, ;
		Name 		= "oleTreeview"

	Add Object cmbdrive As ComboBox With ;
		RowSourceType 	= 3, ;
		RowSource 		= "Select cDrive from qDisc Into Cursor qTempDrive", ;
		Height 			= 23, ;
		Left 			= 382, ;
		Style 			= 2, ;
		Top 			= 426, ;
		Width 			= 73, ;
		Name 			= "cmbDrive"

	Add Object btncancelar As CommandButton With ;
		Top 			= 424, ;
		Left 			= 530, ;
		Height 			= 27, ;
		Width 			= 73, ;
		FontName 		= "Tahoma", ;
		FontSize 		= 9, ;
		Picture 		= "", ;
		Caption 		= "\<Cancelar", ;
		PicturePosition = 0, ;
		BackColor 		= Rgb(255,255,255), ;
		Themes 			= .T., ;
		Name 			= "btnCancelar"

	Add Object btnaceptar As CommandButton With ;
		Top 			= 424, ;
		Left 			= 457, ;
		Height 			= 27, ;
		Width 			= 73, ;
		FontName 		= "Tahoma", ;
		FontSize 		= 9, ;
		Picture 		= "", ;
		Caption 		= "\<Aceptar", ;
		PicturePosition = 0, ;
		BackColor 		= Rgb(255,255,255), ;
		Themes 			= .T., ;
		Name 			= "btnAceptar"

	Add Object txtcurrentpath As TextBox With ;
		Height 				= 23, ;
		Left 				= 12, ;
		ReadOnly 			= .T., ;
		Top 				= 426, ;
		Width 				= 365, ;
		DisabledBackColor 	= Rgb(255,255,255), ;
		Name 				= "txtCurrentPath"

	Add Object label1 As Label With ;
		AutoSize 	= .T., ;
		FontBold 	= .T., ;
		FontName 	= "Tahoma", ;
		FontSize 	= 8, ;
		Alignment 	= 0, ;
		BackStyle 	= 0, ;
		Caption 	= "Ruta de Conexión a la Data:", ;
		Height 		= 15, ;
		Left 		= 12, ;
		Top 		= 408, ;
		Width 		= 157, ;
		ForeColor 	= Rgb(106,106,106), ;
		BackColor 	= Rgb(235,235,235), ;
		Style 		= 0, ;
		Name 		= "Label1"

	Add Object oleimageslist As OleControl With ;
		oleclass	= "MSComctlLIB.ImagelistCtrl.2",;
		Top 		= 168, ;
		Left 		= 24, ;
		Height 		= 100, ;
		Width 		= 100, ;
		Name 		= "oleImagesList"

	Add Object label2 As Label With ;
		AutoSize 	= .T., ;
		FontBold 	= .T., ;
		FontName 	= "Tahoma", ;
		FontSize 	= 8, ;
		Alignment 	= 0, ;
		BackStyle 	= 0, ;
		Caption 	= "Unidad", ;
		Height 		= 15, ;
		Left 		= 382, ;
		Top 		= 408, ;
		Width 		= 41, ;
		ForeColor 	= Rgb(106,106,106), ;
		BackColor 	= Rgb(235,235,235), ;
		Style 		= 0, ;
		Name 		= "Label2"

	Add Object lblText As Label With ;
		AutoSize 	= .T., ;
		FontBold 	= .T., ;
		FontName 	= "Tahoma", ;
		FontSize 	= 8, ;
		BackStyle 	= 0, ;
		Caption 	= "Prueba", ;
		Left 		= 12, ;
		Top 		= 10, ;
		Name 		= "lblText"

	Procedure filltree
		Parameters m.path, m.nlevel, m.nCount
		Local DirArr,i,nTotDir,lvl,pkey
		m.path = Alltrim(m.path)
		If Parameters()<2 Or Type("m.nlevel") #"N"
			lvl = 0
		Else &&Parameters()<2 Or Type("m.nlevel") #"N"
			lvl = m.nlevel
		Endif &&Parameters()<2 Or Type("m.nlevel") #"N"
		If Parameters()<2 Or Type("m.nCount") #"N"
			Cnt = 0
		Else &&Parameters()<2 Or Type("m.nCount") #"N"
			Cnt = m.nCount
		Endif &&Parameters()<2 Or Type("m.nCount") #"N"
		lvl 	= lvl + 1
		Cnt 	= Cnt + 1
		pkey 	= Lower(Substr(m.path,1,Rat("\",m.path,2)))+"_"
* Add items to treeview control
		o = This.oletreeview
		If Cnt = 1
			oNode 		= o.nodes.Add(,1,Lower(m.path)+"_",Lower(m.path),,)
			oNode.Image = 1         &&"Folder"
		Else &&Cnt = 1
			oNode 		= o.nodes.Add(m.pkey,4,Lower(m.path)+"_",Lower(m.path),,)
			oNode.Image	= 2
		Endif &&Cnt = 1
		Dimension DirArr[1,1]
		nTotDir = Adir(DirArr,m.path+"*.","D")
		Asort(DirArr)
		For i = 1 To m.nTotDir
			If DirArr[m.i,1] != '.' And Atc('D',DirArr[m.i,5])#0
				This.filltree(m.path+DirArr[m.i,1]+'\', m.lvl, m.cnt)
			Endif
		Endfor &&i = 1 To m.nTotDir
	Endproc

	Procedure filldrive
		For i = 65 To 90 Step 1
			cVol = Chr(i) + ":\"
			If Directory(cVol)
				Insert Into qDisc (cDrive) Values(cVol)
			Else &&DIRECTORY(cVol)
			Endif &&DIRECTORY(cVol)
		Endfor
	Endproc

	Procedure Init
		Lparameters tcDirectory As String, tcText As String, tcCaption As String

		If Type("THIS.oleTreeview") # "O" Or Isnull(This.oletreeview)
			Return .F.
		Else &&TYPE("THIS.oleTreeview") # "O" OR ISNULL(THIS.oleTreeview)
		Endif &&TYPE("THIS.oleTreeview") # "O" OR ISNULL(THIS.oleTreeview)

		If Type("THIS.oleImagesList") # "O" Or Isnull(This.oleimageslist)
			Return .F.
		Else &&TYPE("THIS.oleImagesList") # "O" OR ISNULL(THIS.oleImagesList)
		Endif &&TYPE("THIS.oleImagesList") # "O" OR ISNULL(THIS.oleImagesList)

		If Empty(tcCaption)
			tcCaption = "Seleccione Directorio"
		Else &&EMPTY(tcCaption)
		Endif &&EMPTY(tcCaption)

		This.Caption = tcCaption

		This.lblText.Visible = .F.
		If !Empty(tcText)
			This.lblText.Caption = tcText
			This.lblText.Visible = .T.
		Else &&!EMPTY(tcText)
		Endif &&!EMPTY(tcText)

		If File(Addbs(Sys(2023)) + "SAVE.BMP")
			This.btnaceptar.Picture = Addbs(Sys(2023)) + "SAVE.BMP"
		Else &&File(Addbs(SYS(2023)) + "SAVE.BMP")
		Endif &&File(Addbs(SYS(2023)) + "SAVE.BMP")

		If File(Addbs(Sys(2023)) + "DOOR.BMP")
			This.btncancelar.Picture = Addbs(Sys(2023)) + "DOOR.BMP"
		Else &&File(Addbs(SYS(2023)) + "DOOR.BMP")
		Endif &&File(Addbs(SYS(2023)) + "DOOR.BMP")

		This.oleimageslist.ListImages.Add(,"Folder",LoadPicture(Addbs(Sys(2023)) + "FOLDER_OPEN.BMP"))
		This.oleimageslist.ListImages.Add(,"ClosedFolder",LoadPicture(Addbs(Sys(2023)) + "FOLDER_CLOSE.BMP"))
		This.oletreeview.ImageList = This.oleimageslist.Object

		If !Empty(tcDirectory) And Directory(tcDirectory)
			Thisform.txtcurrentpath.Value = ADDBS(tcDirectory)
			Thisform.oletreeview.nodes.Clear
			Thisform.filltree(ADDBS(tcDirectory))
			Thisform.cmbdrive.Value = Upper(Left(tcDirectory,3))
		Else &&!EMPTY(tcDirectory) AND DIRECTORY(tcDirectory)
		Endif &&!EMPTY(tcDirectory) AND DIRECTORY(tcDirectory)

	Endproc

	Procedure Load
		Set Safety Off
		Create Cursor qDisc(cDrive c(3))
		This.filldrive()
		This.LoadImages()
	Endproc

	Procedure LoadImages
		local imgSave, imgExit, imgFolderClose, imgFolderOpen
		text to m.imgSave noshow
		Qk02AwAAAAAAADYAAAAoAAAAEAAAABAAAAABABgAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAAeGZfYEpCnpGK7ezn7ezn7ezn7ezn7ezn7ezn7ezn7ezn7ezn7eznbVlSYEpCeGZfeGZfYEpCnpGK7ezn7ezn7ezn7ezn7ezn7ezn7ezn7ezn7ezn7eznbVlSYEpCeGZfeGZfYEpCnpGK7ezn7ezn7ezn7ezn7ezn7ezn7ezn7ezn7ezn7eznbVlSYEpCeGZfeGZfYEpCZqKyI73rI73rI73rI73rI73rI73rI73rI73rI73rI73rZWFfYEpCeGZfeGZfYEpCYqO1GLvrqOX4t+r5t+r5muH2W8/xFrrrFrrrFrrrFrrrZWJgYEpCeGZfeGZfYEpCYqO1GLvrkN72muH2gtr0FrrrFrrrFrrrFrrrFrrrFrrrZWJgYEpCeGZfeGZfYEpCYqO1FrrrLsHtM8PuKsDtFrrrFrrrFrrrFrrrFrrrFrrrZWJgYEpCeGZfeGZfYEpCY1xYZ3FyZ3FyZ3FyZ3FyZ3FyZ3FyZ3FyZ3FyZ3FyZ3FyYU1GYEpCeGZfeGZfYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCeGZfeGZfYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCYEpCeGZfeGZfYEpCYEpCp56Yt7Grt7Grt7Grt7Grl42HjYJ8jYJ8jYJ8ZE9HYEpCYEpCeGZfeGZfYEpCYEpC4N/a7ezn7ezn7ezn6OfjyMvIm5SPmpKNxMbDalZPYEpCYEpCeGZfeGZfYEpCYEpC4N/a7ezn7ezn7ezn1dfTx8rHe2xlfW5nwcK/alZPYEpCYEpCeGZfeGZfYEpCYEpC4N/a7ezn7ezn6OfjyMvIx8rHg3Zvhnp0wcK/alZPYEpCYEpCopWReGZfYEpCYEpC4N/a7ezn7ezn1dfTx8rHx8rHcF5Xb11VwcK/alZPYEpCn5KN/v7+eGZfYEpCYEpC4N/a7ezn6OfjyMvIx8rHx8rHx8rHx8rHx8rHalZPoJOO/v7+////
		Endtext

		text to m.imgExit noshow
		Qk02AwAAAAAAADYAAAAoAAAAEAAAABAAAAABABgAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAA////////////////////+/v/xcX4fX3vUlLp////////////////////////////////////////xMT3b2/tKyvkGxviGxviMzPiz8/Pz8/Pz8/P3d3d////////////////////////WlrqGxviGxviGxviGxviNzfm/Pz8/Pz8/Pz8y8vL////////////////////////WlrqGxviGxviGxviGxviNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviGxviNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviGxviNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviGxviNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviJyfjNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviTEzoOjrn////////////zMzM////////////////////////WlrqGxviGxviGxviGxviNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviGxviNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviGxviNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviGxviNzfm////////////zMzM////////////////////////WlrqGxviGxviGxviGxviNzfm/Pz8/Pz8/Pz8y8vL////////////////////////xMT3b2/tKyvkGxviGxviMzPiz8/Pz8/Pz8/P3d3d////////////////////////////////+/v/xcX4fX3vUlLp////////////////////////////
		Endtext

		text to m.imgFolderClose noshow
		Qk02AwAAAAAAADYAAAAoAAAAEAAAABAAAAABABgAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAA////////////////////////////////////////////////////////////////v+75quj3quj3quj3quj3quj3quj3quj3quj3quj3quj3quj3ren49v3+////////WtHwSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vsur4////////W9DxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vb9jy////////LcHtS87vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7v4/f8////FrrrXtLxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7voOX3////FrrrU83wSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vX9Tx/v//FrrrIL3sT9DwSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7v0fP7FrrrFrrrYdLySs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vjuD1FrrrFrrrRcjvS87vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vYtXxFrrrFrrrFrrrIL3rI77sI77sI77sI77sI77sI77sI77sI77sI77sI77srOf4+f3+FrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrFrrrr+f4////FrrrFrrrFrrrFrrrFrrrFrrrMMLte9j0e9j0e9j0e9j0e9j0e9j0ftn07Pn9////FrrrFrrrFrrrFrrrFrrrGbvrx+76////////////////////////////////////rOb4kd/2kd/2kd/2kd/2suj4////////////////////////////////////////////////////////////////////////////////////////////////////////
		Endtext

		text to m.imgFolderOpen noshow
		Qk02AwAAAAAAADYAAAAoAAAAEAAAABAAAAABABgAAAAAAAAAAADEDgAAxA4AAAAAAAAAAAAA////////////////////////////////////////////////////////////////v+75quj3quj3quj3quj3quj3quj3quj3quj3quj3quj3quj3ren49v3+////////W9LxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vsur4////////XNDxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vb9jy////////QcfuTM7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7v4/f8////QMbukt/zSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7voOX3////QMbu0O72Ss7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vX9Tx/v//QMbu3e/0V9LxSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7v0fP7QMbu3O/0p+T0Ss7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vjuD1QMbu3O/02fD1TM7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vSs7vYtXxKr/su9jfzuPo2evw0OTp1unu0uXq1Ojt1Ojt0uXq1unu0OTp2ezx3fD15fT4+f3+FrrrmdHiyN/k3O/03O/03O/03O/03O/03O/03O/03O/03O/03O/03O/06fX4////Frrrks3fhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/Rhr/R9/v8////FrrrKb7qNMDrNMDrNMDrN8Dqzu739Pn79Pn79Pn79Pn79Pn79Pn79Pn7////////rOb4kd/2kd/2kd/2kd/2s+n4////////////////////////////////////////////////////////////////////////////////////////////////////////
		Endtext

		=Strtofile(Strconv(m.imgSave,14),Addbs(SYS(2023)) + "SAVE.BMP")
		=Strtofile(Strconv(m.imgExit,14),Addbs(SYS(2023)) + "DOOR.BMP")
		=Strtofile(Strconv(m.imgFolderClose,14),Addbs(SYS(2023)) + "FOLDER_CLOSE.BMP")
		=Strtofile(Strconv(m.imgFolderOpen,14),Addbs(SYS(2023)) + "FOLDER_OPEN.BMP")
	Endproc

	Procedure Destroy
		Try
			cFile = Addbs(Sys(2023)) + "SAVE.BMP"
			Delete File &cFile
		Catch
		Endtry

		Try
			cFile = Addbs(Sys(2023)) + "DOOR.BMP"
			Delete File &cFile
		Catch
		Endtry

		Try
			Delete File (Addbs(Sys(2023)) + "FOLDER_CLOSE.BMP")
			Delete File &cFile
		Catch
		Endtry

		Try
			cFile = Addbs(Sys(2023)) + "FOLDER_OPEN.BMP"
			Delete File &cFile
		Catch
		Endtry

	Endproc

	Procedure oletreeview.NodeClick
*** ActiveX Control Event ***
		Lparameters Node
		For Each loNode In Thisform.oletreeview.nodes
			loNode.Image=2
			Node.Image=1
		Endfor
		Thisform.txtcurrentpath.Value = Node.Text
	Endproc

	Procedure cmbdrive.InteractiveChange
		Thisform.oletreeview.nodes.Clear
		Thisform.filltree(This.Value)
	Endproc

	Procedure btncancelar.Click
		Thisform.cPath = ""
		Thisform.Hide()
	Endproc

	Procedure btnaceptar.Click
		Thisform.cPath = ""
		If !Empty(Thisform.txtcurrentpath.Value)
			Thisform.cPath = Thisform.txtcurrentpath.Value
		Else &&!EMPTY(THISFORM.txtCurrentPath.VALUE)
		Endif &&!EMPTY(THISFORM.txtCurrentPath.VALUE)
		Thisform.Hide()
	Endproc

	Procedure oleimageslist.Init
		With This
			.ImageHeight = 16
			.ImageWidth  = 16
		Endwith
	Endproc
Enddefine
