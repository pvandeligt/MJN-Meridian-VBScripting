'**************************************************************EVENTS*********************************************************
'********************************************************************************************************************************
Dim strTmpPath


Sub DocCopyMoveEvent_BeforeCopy(Batch, TargetFolder)
	'WinMsgBox TargetFolder.Name
    'WinMsgBox TargetFolder.Property("Project.Projectomschrijving")
	If Not Document Is Nothing Then
     If client.ImportDetails = AS_ID_CREATEPROJCOPY Or  client.ImportDetails = AS_ID_CREATEPROJCOPYWLOCK  Then
     	Document.StatusText = "Locked in: " & TargetFolder.Name & " " & TargetFolder.Property("Project.Projectomschrijving")
     End If
     If client.ImportDetails = AS_ID_CREATEPROJCOPY Or  client.ImportDetails = AS_ID_CREATEPROJCOPYWLOCK  Then
     	If  Document.StatusText = "Unchanged" Then
    		Document.ChangeWorkflowState AS_WF_RELEASED , "Document Control", ""
    	End If
           
        If ucase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
            Document.Project_Project_Root = "Master"
        End If
        
        If ucase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
    		'RAC 08-11-2012 Folderstructuur overnemen van Master in Project
        	If Document.ProcessCell = "PC90 Project" Then
                Document.Project_Project_Root = "Project"
        	Else
                Document.Project_Project_Root = "Reference"
        	End If
    	End If
        If ucase(Document.DocumentType) = "OBJECTS" Then
            Document.Project_Project_Root = "Object"
        End If
     End If
	ElseIf Not Folder Is Nothing Then

	End If
End Sub


Sub DocGenericEvent_BeforeNewDocument(Batch, Action, SourceFile, DocType, DocTemplate)
	'2017-06-19, RC, Als een document erop gesleept wordt dan 
    If Not Folder Is Nothing Then
    	'2017 06 29 RC, Alleen gegevens ophalen als er in de projectbranch gesleept wordt.
    	If Document.Branch = "Project" Then
    	 If Document.Projects = "" Then
	    	Document.Projects = Folder.Project_ProjectAndDescription
            '2017-06-20 RC, als een document lager geimporteerd wordt de waarde ophalen van het project
            If Document.Projects = "" Then
            	If Folder.ParentProject.Project_ProjectAndDescription <> "" Then
                	Document.Projects = Folder.ParentProject.Project_ProjectAndDescription
                End If
            End If
	     End If
        End If
	End If

	'als documenten gereleased worden naar de master de juiste branch verplaatsen
    If client.ImportDetails = AS_ID_RELEASETOMASTER Then
    	If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
        	'WinMsgBox "Set Branch Master"
    		Document.Branch = "Master"
        End If
        If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
        	'WinMsgBox "Set Branch Reference"
            Document.Branch = "Reference"
        End If
        If ucase(Document.DocumentType) = "OBJECTS" Then
        	'WinMsgBox "Set Branch Reference"
    		Document.Branch = "Object"
        End If
    End If
    
    'RAC 01-12-2016, Reference documenten direct in Reference plaatsen
    If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
    	 'RAC 24-03-2017 branch alleen vullen als deze niet project is
    		If Document.Branch <> "Project" Then
            	Document.Branch = "Reference"
            End If
    End If
  
    
    If client.ImportDetails <> AS_ID_RELEASETOMASTER Then
		'controleren of nieuwe documenten aangemaakt worden op de projectfolder, anders nieuw doc afbreken    
		If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Or UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
        		Document.Site = "NIJ"
                
    		    'This code shows how you can initialize the property values For new documents.
    			bInitializeProperties = False
    			' You can (for example) limit this to Importing and Batchscan
    			If (Action = AS_IT_IMPORTED) Or (Action = AS_IT_BATCHSCAN) Then
    				' Or any other criteria
        			bInitializeProperties = DoYouWantToInitializeThePropertyValues(DocType)
    			End If
          
    			If bInitializeProperties Then
        			' Here we initialize the document properties 
        			' This saves the user from having to type these value in the wizard
        			InitializeDocumentProperty Batch, Document, DocType
    			End If
    
    			'RAC 22-02-2013 als een master document wordt geimporteerd en het process is bekend, de eerste 5 karakters invoeren bij title2
    			If Document.Branch = "Master" Then
    				If Document.AsBuilt_Process <> "" Then
        				If Document.TitleBlockLine2 = "" Then
        					Document.TitleBlockLine2 = Left(Document.Process, 5)
           				End If
                    	If Document.TitleBlockLine1 = "" Then
        					Document.TitleBlockLine1 = Mid(Document.Process,7,Len(Document.Process)-6)
                    	End If
        			End If
    			End If    
    	Else
    		'afhandeling van Objecten
    		Document.Site = "NIJ"    
    	End If
	End If    
    If ucase(Document.DocumentType) = "OBJECTS" Then
    	If Document.Object_Status = "" Then
        	Document.Object_Status = "New"
        End If
        ' RAC 19 sept 2016 voor nieuwe objecten de branch direct goed zetten zodat unit / EM goed gevuld worden
        Document.Branch = "Object"
    Else
    	Document.ForceName = True
    End If   
    
    '07-11-2014 RAC volgnummer ophogen bij batch import
    If client.ImportType = AS_IT_IMPORTED  Then
	    If Not Batch.IsFirstInBatch Then
			Document.DocSequenceNumber = GetSequence
            'WinMsgBox Document.DocSequenceNumber
	    End If
        'title block 1 vullen met bestandsnaam
        Title1 = Document.FileName
        
        Point = InStrRev(Title1 ,".",-1, vbTextCompare)
    
    	Title1 = Left(Title1 ,(Point-1))
    
		Document.TitleBlockLine1 = Title1 
    End If
 
End Sub

Sub DocGenericEvent_AfterNewDocument(Batch, Action, SourceFile, DocType, DocTemplate)
	If Action <> AS_IT_MOVED Then 
    	If Document.DocumentType.DisplayName = "Objects" Then
    		Document.Branch = "Object"
    	End If
    End If
    
        
    'RAC, 01-12-2016, General documenten direct in reference opslaan    
    'If ucase(Document.DocumentType) <> "GENERAL DOCUMENTS" Then
    'RAC, 24-03-2016 General op project gesleept wel in project opslaan
    If ucase(Document.DocumentType) <> "GENERAL DOCUMENTS" Or (ucase(Document.DocumentType) = "GENERAL DOCUMENTS" And Document.Branch = "Project") Then
      If client.ImportType = AS_IT_IMPORTED  Then
    	 If Not Batch.IsFirstInBatch Then
         	Document.Project_Project = Batch.Argument ("Project_Project")
        	Document.Project_Projectnummer = Batch.Argument ("Project_Projectnummer")
        	Document.Projects  = Batch.Argument ("Project_Projectnummer")
        	Document.Project_Projectomschrijving = Batch.Argument ("Project_Projectomschrijving")
        	Document.Project_Projectleider  = Batch.Argument ("Project_Projectleider") 
        	Document.Project_Comments  = Batch.Argument ("Project_Comments")
        	Document.Project_Date_Inquiry = Batch.Argument ("Project_Date_Inquiry") 
        	Document.Project_Date_Req = Batch.Argument ("Project_Date_Req")
        	Document.Project_EndDate = Batch.Argument ("Project_EndDate")
        	Document.Project_StartDate = Batch.Argument ("Project_StartDate")
        	Document.Project_WorkOrderNumber = Batch.Argument ("Project_WorkOrderNumber")
            
        End If
     End If
     If Document.Branch = "Project" Then
    	Document.Project_ProjectAndDescription = Document.Project_Projectnummer & " " & Document.Project_Projectomschrijving
     End If
     
    End If
    
'If ucase(vault.User) <> "RAC" Then    
    'Alleen nieuwe documenten verplaatsen naar project   
    If client.ImportDetails <> AS_ID_RELEASETOMASTER And client.ImportDetails <> AS_ID_CREATEPROJCOPY And  client.ImportDetails <> AS_ID_CREATEPROJCOPYWLOCK  Then
    	'RAC, 01-12-2016, General documenten direct in reference opslaan  
        'If ucase(Document.DocumentType) <> "GENERAL DOCUMENTS" Then  
        'RAC, 24-03-2016 General op project gesleept wel in project opslaan
	    If ucase(Document.DocumentType) <> "GENERAL DOCUMENTS" Or (ucase(Document.DocumentType) = "GENERAL DOCUMENTS" And Document.Branch = "Project") Then
         If Batch.IsFirstInBatch Then
        	Project = Split(Document.Projects," ")
        	Batch.Argument("Project") = Project(0)
        	Document.Project_Projectnummer = Project(0)
            Batch.Argument("Projectomschrijving") = Mid(Document.Projects,Len(Project(0))+1)
         End If
         Document.Branch = "Project"
        End If
        
        If ucase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
    		'RAC 08-11-2012 Folderstructuur overnemen van Master in Project
        	NewFolder = "Project\NIJ\" & Batch.Argument("Project")  & "\Master\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Discipline & "\" & Document.DisciplineClass
        	Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
            Document.Project_Project_Root = "Master"
    	End If
        
    	If ucase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
    		'RAC 08-11-2012 Folderstructuur overnemen van Master in Project
        	If Document.ProcessCell = "PC90 Project" Then
        		NewFolder = "Project\NIJ\" & Batch.Argument("Project")  & "\Project\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Discipline & "\" & Document.DisciplineClass
                Document.Project_Project_Root = "Project"
                Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
            'RAC 01-12-2016, uitgezet, reference documenten behalve PC90 direct in reference plaatsen
            'RAC 24-03-2016 weer aangezet voor reference documenten die op project gesleept zijn
        	Else
            	If Document.Branch = "Project" Then
                 NewFolder = "Project\NIJ\" & Batch.Argument("Project")  & "\Reference\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Discipline & "\" & Document.DisciplineClass
                 Document.Project_Project_Root = "Reference"
                 Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
                End If
        	End If
    	End If
        If ucase(Document.DocumentType) = "OBJECTS" Then
	        If client.ImportDetails = AS_ID_CREATEPROJCOPY Or  client.ImportDetails = AS_ID_CREATEPROJCOPYWLOCK  Then
    			NewFolder = Document.ParentFolder.Path  & "\Object\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Object_Unit & "\" & Document.Object_EquipmentModule 
        		Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
                Document.Project_Project_Root = "Object"
        	Else
        		NewFolder = "Project\NIJ\" & Batch.Argument("Project")  & "\Object\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Object_Unit & "\" & Document.Object_EquipmentModule 
        		Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
                Document.Project_Project_Root = "Object"
        	End If
        End If
        If Not Document.MasterDocument Is Nothing Then
        	SetProjectCopy("Locked in: "& Batch.Argument("Project") & " " & Batch.Argument("Projectomschrijving"))  
        End If
        
        
    End If
    
    
    If client.ImportDetails <> AS_ID_RELEASETOMASTER Then
    	'RAC, 01-12-2016, General documenten direct in reference opslaan  
        'If ucase(Document.DocumentType) <> "GENERAL DOCUMENTS" Then  
       'RAC, 24-03-2016 General op project gesleept wel in project opslaan
	   If ucase(Document.DocumentType) <> "GENERAL DOCUMENTS" Or (ucase(Document.DocumentType) = "GENERAL DOCUMENTS" And Document.Branch = "Project") Then
    	'WinMsgBox Document.ParentProject.Property("Project_Project")
    	Document.Project_Projectnummer  = Document.ParentProject.Property("Project_Project")
        'RAC 0309 added
        Document.Projects  = Document.ParentProject.Property("Project_Project")
        '--
        Document.Project_Projectomschrijving = Document.ParentProject.Property("Project_Projectomschrijving")
        
        Document.Project_ProjectAndDescription = Document.Project_Projectnummer & " " & Document.Project_Projectomschrijving
        Document.Project_Projectleider  = Document.ParentProject.Property("Project_Projectleider")
        Document.Project_Comments  = Document.ParentProject.Property("Project_Comments")
        Document.Project_Date_Inquiry = Document.ParentProject.Property("Project_Date_Inquiry")
        Document.Project_Date_Req = Document.ParentProject.Property("Project_Date_Req")
        Document.Project_EndDate = Document.ParentProject.Property("Project_EndDate")
        Document.Project_StartDate = Document.ParentProject.Property("Project_StartDate")
        Document.Project_WorkOrderNumber = Document.ParentProject.Property("Project_WorkOrderNumber")
       End If
    End If
    
        
    If client.ImportDetails = AS_ID_CREATEPROJCOPY Or  client.ImportDetails = AS_ID_CREATEPROJCOPYWLOCK  Then
        'RAC 08-11-2012 Document releasen als er een project copy van is gemaakt
        'WinMsgBox "Create project copy"
    	If  Document.StatusText = "Unchanged" Then
    		Document.ChangeWorkflowState AS_WF_RELEASED , "Document Control", ""
    	End If
           
        If ucase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
    		'RAC 08-11-2012 Folderstructuur overnemen van Master in Project
            'WinMsgBox Document.ParentFolder.Path
        	NewFolder = Left(Document.ParentFolder.Path,19) &  "\Master\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Discipline & "\" & Document.DisciplineClass
            Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
            Document.Project_Project_Root = "Master"
        End If
        
        If ucase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
    		'RAC 08-11-2012 Folderstructuur overnemen van Master in Project
        	If Document.ProcessCell = "PC90 Project" Then
        		NewFolder = "Project\NIJ\" & Document.Projects  & "\Project\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Discipline & "\" & Document.DisciplineClass
                Document.Project_Project_Root = "Project"
                Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
            'RAC 01-12-2016, uitgezet, reference documenten behalve PC90 direct in reference plaatsen
        	'Else
            '    NewFolder = "Project\NIJ\" & Document.Projects  & "\Reference\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Discipline & "\" & Document.DisciplineClass
            '    Document.Project_Project_Root = "Reference"
        	End If
    	End If
        If ucase(Document.DocumentType) = "OBJECTS" Then
            'Volantis, RAC 28 sept 2016, Object folder werd verkeerd opgebouwd
           	'NewFolder = Document.ParentFolder.Path  & "\Object\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Object_Unit & "\" & Document.Object_EquipmentModule
            NewFolder = "Project\NIJ\" & Document.Projects  & "\Object\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Object_Unit & "\" & Document.Object_EquipmentModule
            Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
            Document.Project_Project_Root = "Object"
        End If
       
     End If
     
    ' When the new document is created we store the property values that we will use 
    ' to initialize the next document
    If client.ImportDetails <> AS_ID_RELEASETOMASTER Then
    	bInitializeProperties = False
    	' You can (for example) limit this to Importing and batchscan
    	If (Action = AS_IT_IMPORTED) Or (Action = AS_IT_BATCHSCAN) Then
    		' Or any other criteria
        	bInitializeProperties = DoYouWantToInitializeThePropertyValues(DocType)
    	End If
    
    	If bInitializeProperties Then
        	StoreDocumentPropertyForThisBatch Batch, Document, DocType
    	End If
    End If
    
    'RAC, 01-12-2016, General documenten direct in reference opslaan    
    'If ucase(Document.DocumentType) <> "GENERAL DOCUMENTS" 
    'RAC, 24-03-2016 General op project gesleept wel in project opslaan
	If ucase(Document.DocumentType) <> "GENERAL DOCUMENTS" Or (ucase(Document.DocumentType) = "GENERAL DOCUMENTS" And Document.Branch = "Project") Then
     If client.ImportType = AS_IT_IMPORTED  Then
    	If batch.IsFirstInBatch Then
        	Document.Project_Project = Document.ParentProject.Property("Project_Project")
            'WinMsgBox Document.ParentProject.Property("Project_Project")
            Document.Project_Projectnummer = Document.ParentProject.Property("Project_Project")
           	Document.Projects  = Document.ParentProject.Property("Project_Projectnummer")
	       	Document.Project_Projectomschrijving = Document.ParentProject.Property("Project_Projectomschrijving")
        	Document.Project_Projectleider  = Document.ParentProject.Property("Project_Projectleider")
        	Document.Project_Comments  = Document.ParentProject.Property("Project_Comments")
        	Document.Project_Date_Inquiry = Document.ParentProject.Property("Project_Date_Inquiry")
        	Document.Project_Date_Req = Document.ParentProject.Property("Project_Date_Req")
        	Document.Project_EndDate = Document.ParentProject.Property("Project_EndDate")
        	Document.Project_StartDate = Document.ParentProject.Property("Project_StartDate")
        	Document.Project_WorkOrderNumber = Document.ParentProject.Property("Project_WorkOrderNumber")
            
            Batch.Argument ("Project_Project") = Document.ParentProject.Property("Project_Project")
            Batch.Argument ("Project_Projectnummer") = Document.ParentProject.Property("Project_Projectnummer")
        	Batch.Argument ("Project_Projectomschrijving") = Document.ParentProject.Property("Project_Projectomschrijving")
        	Batch.Argument ("Project_Projectleider")  = Document.ParentProject.Property("Project_Projectleider")
        	Batch.Argument ("Project_Comments")  = Document.ParentProject.Property("Project_Comments")
        	Batch.Argument ("Project_Date_Inquiry") = Document.ParentProject.Property("Project_Date_Inquiry")
        	Batch.Argument ("Project_Date_Req") = Document.ParentProject.Property("Project_Date_Req")
        	Batch.Argument ("Project_EndDate") = Document.ParentProject.Property("Project_EndDate")
        	Batch.Argument ("Project_StartDate") = Document.ParentProject.Property("Project_StartDate")
        	Batch.Argument ("Project_WorkOrderNumber") = Document.ParentProject.Property("Project_WorkOrderNumber")
        End If
     End If
    End If
           
    Document.Site = "NIJ"
    If ucase(Document.DocumentType) = "OBJECTS" Then
    		Document.Object_TechnicalIDNumberANDDescription = Document.Object_TechnicalIDNumber & " " & Document.Object_TechnicalIDDescription 
    End If    
'
	If ucase(Document.DocumentType) = "OBJECTS" Then
    	Document.SetModified(Today)
    End If

	'RAC 11-12-2014 Title goed vullen
   	Document.Title = Trim(Document.TitleBlockLine1 & " " & Document.TitleBlockLine2 & " " & Document.TitleBlockLine3 & " " & Document.TitleBlockLine4)
    
    
    'RAC 27-03-2017
    'General documents die in Projects worden gesleept direct retiren en naar het archief verplaatsen
    If (ucase(Document.DocumentType) = "GENERAL DOCUMENTS" And Document.Branch = "Project") Then
    	Document.ReleaseChange
        
        '2017 06 29 RC, alleen PC90 Project documenten archiveren
        If Document.ProcessCell = "PC90 Project" Then
	    	Archive_Document()
   			'retire document
    		Document.ChangeWorkflowState AS_WF_RETIRED, "", ""
        End If
    End If 
   
End Sub


Sub DocProjectCopyEvent_BeforeDiscardFromProject(Batch)

If Not Document.MasterDocument Is Nothing Then

	strTmpPath = Left(Document.Path,19)
	
	'Add your code here
    'WinMsgBox "Discard"
    SetProjectCopy("")
    Document.MasterDocument.Property("ProjectCopy") = ""
    Document.MasterDocument.ApplyPropertyValues 
    
       'status herstellen van Master document to Released
	    Dim Criteria 
        Dim Doc
        TemplatePath = ""
		Criteria = Array(Array("Custom.Branch",  IC_OP_EQUALS, "Master"))
        Teller = Vault.FindDocuments(Document.FileName,,Criteria,False).Count
       
        If Teller = 1 Then	
        	For Each Doc In Vault.FindDocuments(Document.FileName,,Criteria,False)
                Doc.StatusText = "Released"
                Doc.ApplyPropertyValues
 			Next 
        End If
    
    
End If    

End Sub

Sub DocProjectCopyEvent_AfterDiscardFromProject(Batch)
	'lege folders verwijderen
    If Ucase(mid(strTmpPath,2,7)) = "PROJECT" Then
     
    	'lege folders verwijderen
        
        'If Batch.IsLastInBatch Then
    	'	exefile = "\\" & vault.ServerName & "\amm3ext$\" & vault.Name & "\BCFolderCleanup.exe " & """" & strTmpPath & """"
        ' 	cmdline = quote(exefile)
        ' 	Execute ("Set WshShell = CreateObject (""WScript.Shell"")")
    	'	Execute ("ReturnCode = WshShell.Run(cmdline, 1, False)")
        'End If
    
    	'client.Refresh(AS_RF_CHANGED_CURRENTVIEW)
    	'client.Refresh(AS_RF_CHANGED_CURRENTFOLDER)
    End If
End Sub


Sub SetProjectCopy(value)
    ' 
    Dim dr, Doc 
    Dim newdocid
    Set dr = AMCreateObject("AMDocumentRepository", True)
    dr.OpenRepository vault.Name, "", Empty, Nothing, True, Nothing, 1, Nothing
    Set Doc = dr.GetFSObject(Document.MasterDocument.id)
    Call WriteNewProps (dr, Doc, value)
    Set Doc = Nothing
    dr.CloseRepository True
    Set dr = Nothing
    Set objDoc = Nothing
End Sub

Sub WriteNewProps (dr, Doc, Value )
    'used in function to copydoc: writes specific properties
    Dim oColl
    Set oColl = Doc.LoadProperties("Custom")
    oColl.Get("ProjectCopy").Value = Value 
               
    Doc.SaveProperties oColl

    Set oColl = Nothing
End Sub


Sub DocGenericEvent_BeforeCalculateFileName(Batch)
	If Not Document Is Nothing Then
		If ucase(Document.DocumentType) = "OBJECTS" Then
        	If client.ImportDetails <> AS_ID_CREATEPROJCOPY And  client.ImportDetails <> AS_ID_CREATEPROJCOPYWLOCK  Then
    			'Controleren of document niet al bestaat in de kluis.
    			Teller = Vault.FindDocuments(Document.Object_TechnicalIDNumber & ".obj").Count
       			If ( Teller >= 1) Then
                    WinMsgBox "Object is " & Teller & " times found, object is NOT unique"
       				Batch.Abort "Object is " & Teller & " times found, object is NOT unique"
                Else
      				'WinMsgBox "Oké, geen dubbele objecten gevonden."
    			End If
			End If
    	End If
        If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Or UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
        	If client.ImportDetails <> AS_ID_CREATEPROJCOPY And  client.ImportDetails <> AS_ID_CREATEPROJCOPYWLOCK   And client.ImportDetails <> AS_ID_RELEASETOMASTER  Then
    			'Controleren of document niet al bestaat in de kluis.
                'WinMsgBox  Document.NamePart & Document.DocSequenceNumber & Document.NameExt
                
                If Not Document.ForceFileName Then
 					Teller = Vault.FindDocuments(Document.FileName).Count
                    If ( Teller > 1) Then
                   			WinMsgBox "Document is " & Teller & " times found, document is NOT unique, creation of document is aborted"
       						Batch.Abort "Document is " & Teller & " times found, document is NOT unique"
                   	End If               
                Else
                	'WinMsgBox Document.NamePart & Document.DocSequenceNumber
                	Teller = Vault.FindDocuments(Document.NamePart & Document.DocSequenceNumber).Count
       				If ( Teller >= 1) Then
                   		If ( Teller = 1) Then
                           	'03-09-2015 Robin, uitgezet, volgens mij overbodig
                    		'WinMsgBox Document.FileName 
                    		'Newname =  Document.NamePart & Document.DocSequenceNumber & FileExtension(Document.FileName)
                            'WinMsgBox Newname 
                    		'If Document.FileName <> Newname Then
                   			'	WinMsgBox "Document is " & Teller & " times found, document is NOT unique, creation of document is aborted"
       						'	Batch.Abort "Document is " & Teller & " times found, document is NOT unique"
                    		'End If
                   		Else
    						WinMsgBox "Document is " & Teller & " times found, document is NOT unique, creation of document is aborted"
       						Batch.Abort "Document is " & Teller & " times found, document is NOT unique"
                   		End If 
    				End If
            	End If
			End If
        End If
        
        
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If

	
End Sub

Sub DocGenericEvent_BeforeReplaceContent(Batch, SourceFile)
	If Document.Branch = "Master" Then
    	Batch.PrintDetails("It is not allowed to import documents in the Master. Please assign the document to a project first.")
        Batch.Abort("It is not allowed to import documents in the Master. Please assign the document to a project first.")
    End If
	'RAC 07-11-2012 - code toegeveogd zodat een document direct overschreven kan worden (zodat er niet eerst een quick change uitgevoerd hoeft te worden)
    If Document.StatusText = "Released" Then
    	Document.StartChange
    	client.Refresh(AS_RF_CHANGED_PROPERTIES)
    End If
End Sub

Sub DocProjectCopyEvent_BeforeReleaseToMaster(Batch, MasterDoc, ProjectCopyChanged)
	'Add your code here
    If UCase(MasterDoc.DocumentType) = "GENERAL DOCUMENTS" Then
        	If Document.ProcessCell = "PC90 Project" Then
          		Batch.Abort("Not allowed to release project document to Master, user Retire function to Archive document")
            End If
   
    End If
    'Robin 17-06-2015
    If ucase(Document.DocumentType) = "OBJECTS" Then
    	'RAC 24-03-2017 Controleren of room van object niet veranderd is anders afbreken en via OMM laten gaan
        If Trim(Document.Object_Room) <> "" And Document.Object_ObjectType = "Instrument" Then
        	'release afbreken als room niet gevonden kan worden
            CheckRoom = Vault.Query("Room").GetValuesEx("Room","Room = '" & Document.Object_Room & "'" ,,,"Room")
            If Not Isarray(CheckRoom) Then
				Batch.FailCurrent("Room niet gevonden in tabel, release naar master wordt afgebroken")
            End If
        End If
        'End 24-03-2017
    	If Document.ParentProject.Property("Project_AssetControlCreated") Then
        	If Batch.IsFirstInBatch Then
        		Answer = WinMsgBox( "Asset control sheet already created are you sure you want to release objects to the Master?",AS_YesNo , "Meridian" )
        		Batch.Argument ("UserAnswer") = Answer 
    			If Batch.Argument ("UserAnswer") = AS_No Then
    				Batch.Abort 
                End If
        	End If
        End If
    End If
    
    Masterdoc.StatusText = "Released"
    Masterdoc.ApplyPropertyValues
End Sub



Sub DocProjectCopyEvent_AfterReleaseToMaster(Batch, MasterDoc, ProjectCopyChanged)
	'copy all properties to the master exept Branch
    'WinMsgBox "Before release to master"
    'WinMsgBox Document.Branch 
    'WinMsgBox "Masterdoc doctype " & MasterDoc.DocumentType 
    Document.CopyProperties "Custom", MasterDoc, Array("Branch")
    If UCase(MasterDoc.DocumentType) = "CONTROLLED DOCUMENTS" Then
        Document.CopyProperties "AsBuilt", MasterDoc
        MasterDoc.Property("Branch") = "Master"
        MasterDoc.Property("Projects") = vbNullString
    End If
    If UCase(MasterDoc.DocumentType) = "GENERAL DOCUMENTS" Then
     	Document.CopyProperties "Reference", MasterDoc
        MasterDoc.Property("Branch") = "Reference"
    End If
    MasterDoc.ApplyPropertyValues 
        
    'RAC 08-11-2012 Vullen PROPS initial en latest project
    If ProjectCopyChanged Then    
    	'WinMsgBox "Project gegevens invullen"
    	If Trim(Masterdoc.Property("Custom.Initialproject")) <> "" Then
        	'WinMsgBox Document.ParentProject.Name & " latest"
    		Masterdoc.Property("Custom.Latestproject") = Document.ParentProject.Name 
        	Masterdoc.ApplyPropertyValues 
     	Else
         	'WinMsgBox Document.ParentProject.Name & " intial"
    		Masterdoc.Property("Custom.Initialproject") = Document.ParentProject.Name 
        	Masterdoc.ApplyPropertyValues 
     	End If
    Else
    	'WinMsgBox "Masterdocument is niet gewijzigd" 
    End If
    
    
    'Lege folders van project verwijderen    
    'Path ophalen
    Dim strPath
    
    blnRoom = False
    strPath =  Left(Document.Path,19)
        
    'document archiveren
    'Object niet archiveren maar alleen release to master en verwijderen
    If ucase(Document.DocumentType) = "OBJECTS" Then
    	If Document.Object_Status = "For Deletion" Then
        
        	'RCL 2021-09-09 Delete record from table
            If ucase(Document.DocumentType) = "OBJECTS" And ucase(Document.Object_ObjectType) = "LOCATION" Then
				Call Vault.Query("Room").DeleteValues(Array("Location", "Room"), Array(Document.Object_Location, Document.Object_TechnicalIDNumberANDDescription))
			End If
        
        	'Object archiveren naar archief
            Document.CopyProperties "Object", MasterDoc
            MasterDoc.Property("Branch") = "Object"
            MasterDoc.Property("Object_Status") = "Archived"
            MasterDoc.Property("Object_ObjectType") = "Archived"
            'Archive master document
            Splits = Split(MasterDoc.Path,"\")
        	i = 1
        	Newfolder = "Archive\"
        	Do While i < UBound(Splits)
        		Newfolder = NewFolder & Splits(i) & "\"
            	i = i + 1
        	Loop
        	
        	NewFolder = Left(NewFolder, Len(NewFolder)-1)
        	MasterDoc.ArchiveSet_ModificationDate = ModificationDate
            MasterDoc.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
            MasterDoc.ChangeWorkflowState AS_WF_RETIRED, "Archived by project: " & Document.Project_Projectnummer, ""
           
            
            'Delete project document
            Document.Delete
        Else
        	Document.CopyProperties "Object", MasterDoc
        	MasterDoc.Property("Branch") = "Object"
        	'WinMsgBox "Doc " & Document.Object_Status
        	'WinMsgBox "MasterDoc " & MasterDoc.Object_Status
        	MasterDoc.Property("Object_Status") = "AsBuilt"
        	MasterDoc.ApplyPropertyValues 
            'OXN Object toevoegen aan table
			If ucase(Masterdoc.DocumentType) = "OBJECTS" And ucase(Masterdoc.Object_ObjectType) = "LOCATION"  Then
            	CheckRoom = Vault.Query("Room").GetValuesEx("Room","Room = '" & Masterdoc.Object_TechnicalIDNumber & "'" ,,,"Room")
            	If Not Isarray(CheckRoom) Then
					Call Vault.Query("Room").AddValues(Array("Location", "Room", "Oldroom"), Array(Masterdoc.Object_Location, Masterdoc.Object_TechnicalIDNumber & " " & Masterdoc.Object_TechnicalIDDescription, "New"))
                End If
			End If
    		Document.Delete()
        End If
    Else
    	If UCase(MasterDoc.DocumentType) = "GENERAL DOCUMENTS" Then
        	If Document.ProcessCell = "PC90 Project" Then
            	Archive_Document()
        		'retire document
    			Document.ChangeWorkflowState AS_WF_RETIRED, "", ""
                'masterdoc verwijderen
'                MasterDoc.Delete
                
                
            Else
            	Document.Delete()
            End If
        Else
	    	Archive_Document()
        	'retire document
    		Document.ChangeWorkflowState AS_WF_RETIRED, "", ""
        End If
    End If
    
    'RAC 21-04-2015 Als documentnaam gewijzigd is in project deze overnemen in Master
    Masterdoc.FileName = Document.FileName
    
    If Ucase(mid(strpath,2,7)) = "PROJECT" Then
    	If Batch.IsLastInBatch Then
			'lege folders verwijderen
    		'exefile = "\\" & vault.ServerName & "\amm3ext$\" & vault.Name & "\BCFolderCleanup.exe " & """" & strPath  & """"
        	'cmdline = quote(exefile)
    		'Execute ("Set WshShell = CreateObject (""WScript.Shell"")")
    		'Execute ("ReturnCode = WshShell.Run(cmdline, 1, False)")
            'client.Refresh(AS_RF_CHANGED_CURRENTVIEW)
	    	'client.Refresh(AS_RF_CHANGED_CURRENTFOLDER)
    	End If
	End If
    If UCase(MasterDoc.DocumentType) = "CONTROLLED DOCUMENTS" Then
		masterdoc.UpdateRendition
    End If
    If UCase(MasterDoc.DocumentType) = "GENERAL DOCUMENTS" Then
		masterdoc.UpdateRendition
    End If
    Client.Goto(MasterDoc)
End Sub

Sub DocWorkflowEvent_BeforeChangeWFState(Batch, SourceState, TargetState, Person, Comment)
	If TargetState = AS_WF_RETIRED Then
       	If ucase(Document.DocumentType) = "OBJECTS" Then
        	'objecten mogen niet geretired worden
            If Document.Branch = "Project" Then
            	WinMsgBox "Object is not allowed to be retired in a Project, it is only allowed in the Object structure"
                Batch.Abort 
            End If
    	Else
    		If Document.HasIncomingReferences Then
    			Answer = WinMsgBox( "The document has incoming references, are you sure you want to archive it?",AS_YesNo , "Meridian" )
        		Batch.Argument ("UserAnswer") = Answer 
    			If Batch.Argument ("UserAnswer") = AS_No Then
    				Batch.Abort 
        		End If
     		End If
    	End If
    End If
End Sub



Sub DocWorkflowEvent_AfterChangeWFState(Batch, SourceState, TargetState, Person, Comment)
	'OXN Object verwijderen uit table
	If TargetState = AS_WF_RETIRED Then
		If ucase(Document.DocumentType) = "OBJECTS" And ucase(Document.Branch) = "OBJECT" And ucase(Document.Object_ObjectType) = "LOCATION" Then
			Call Vault.Query("Room").DeleteValues(Array("Location", "Room"), Array(Document.Object_Location, Document.Object_TechnicalIDNumberANDDescription))
		End If
	End If
	'WinMsgBox TargetState
    If TargetState = AS_WF_RETIRED Then
    	If Document.Branch <> "Archive" Then
	    	Archive_Document()
        End If
        If Document.ProcessCell = "PC90 Project" Then
        	Document.InitialProject = Document.Projects
            Document.LatestProject = Document.Projects
        End If
	End If
    If (SourceState = AS_WF_UNCHANGED) And (TargetState = AS_WF_RELEASED) Then
		Document.UpdateRendition
    End If
End Sub

Sub ProjectWorkflowEvent_BeforeExpandItem(SubItems) 
		'WinMsgBox Client.ImportDetails 
		  Select Case Client.ImportDetails 
            Case AS_ID_CREATEPROJCOPY,AS_ID_CREATEPROJCOPYWLOCK,AS_ID_NODETAILS 
        		If Folder.Path = "\" Then 
            		iLevel = 0 
    			Else 
            		iLevel = UBound (Split (Folder.Path, "\")) 
         		End If 
      			'winmsgbox "Level: " & iLevel &vbCrLf & "FolderName: " & Folder.Path & vbCrLf & "ClientDetails: " & Client.ImportDetails 
       			Select Case iLevel 
            		Case 0 
                		'DebugMessage "Case 0"         
                    	For i = LBound (SubItems) To UBound (SubItems) 
                            If SubItems (i)(1) = "Project" Then 
                        		SubItems (i)(2) = AS_PRJITEM_MODE_VISIBLE + AS_PRJITEM_MODE_EXPANDABLE 
                			Else 
                        		SubItems (i)(2) = AS_PRJITEM_MODE_NONE 
                			End If 
            			Next 
            		Case 1 
                		'DebugMessage "Case 1" 
                    	For i = LBound (SubItems) To UBound (SubItems) 
                    		SubItems (i)(2) = AS_PRJITEM_MODE_VISIBLE + AS_PRJITEM_MODE_EXPANDABLE 
                    	Next         
        			Case 2 
                		'DebugMessage "Case 2" 
                    	For i = LBound (SubItems) To UBound (SubItems) 
                        	SubItems (i)(2) = AS_PRJITEM_MODE_VISIBLE + AS_PRJITEM_MODE_SELECTABLE 
            			Next   
        			Case 3 
                		'DebugMessage "Case 3" 
            			For i = LBound (SubItems) To UBound (SubItems) 
                           SubItems (i)(2) = AS_PRJITEM_MODE_NONE 
                        Next 
        			Case 4 
                		'DebugMessage "Case 4" 
            			For i = LBound (SubItems) To UBound (SubItems) 
                    		SubItems (i)(2) = AS_PRJITEM_MODE_NONE 
            			Next 
        			Case Else 
                		For i = LBound (SubItems) To UBound (SubItems) 
                			SubItems (i)(2) = AS_PRJITEM_MODE_VISIBLE + AS_PRJITEM_MODE_EXPANDABLE + AS_PRJITEM_MODE_SELECTABLE 
            			Next 
    			End Select 
    	End Select 
      
End Sub 
Sub DocGenericEvent_OnProperties(Command, Abort)
	'Objecten mogen alleen maar bewerkt worden in Quick Change mode
	'Alle documenten mogen niet gewijzigd worden als deze retired zijn
	If Not Document Is Nothing Then
		'Add your code for document objects
        If Command = AS_PS_CMD_EDIT Then
        	If Document.WorkFlowState = AS_WF_RETIRED  Then
        		WinMsgBox "Editing of properties only permitted when document is in Quick Change"
            	Abort = True
        	Else
        		If ucase(Document.DocumentType) = "OBJECTS" Then
        			If Document.WorkFlowState = AS_WF_RELEASED Then
                    	WinMsgBox "Editing of properties only permitted when document is in Quick Change"
                		Abort = True
                	End If
        		End If
        	End If
        End If
    If Command = AS_PS_CMD_APPLY Then
    	'RAC 11-12-2014 Title goed vullen
        Document.Title = Trim(Document.TitleBlockLine1 & " " & Document.TitleBlockLine2 & " " & Document.TitleBlockLine3 & " " & Document.TitleBlockLine4)
        'RAC 21-04-2015 When apply, rename object according to TechnicalIDNumber
        If ucase(Document.DocumentType) = "OBJECTS" Then
        	Document.FileName = Document.Object_TechnicalIDNumber & ".obj"
		End If
   End If    
        
    ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Sub DocWorkingCopyEvent_BeforeSubmitWC(Batch)
	'Als een VSD document released wordt dan de PDF als rendition opslaan
    If Right(Ucase(Document.FileName),3) = "VSD" Then
    	Document.SaveToFile "c:\temp\" & Document.FileName 
             
    	Set objVisio = CreateObject("Visio.Application")
    	Set objDoc = objVisio.Documents.OpenEx("c:\temp\" & Document.FileName, &H2 + &H10 + &H40)    
    
    	PDFName = Document.FileName
    	PDFName = Left(PDFName,Len(PDFName)-3) & "pdf"
     
    	objDoc.ExportAsFixedFormat 1, "c:\temp\" & PDFName, 1, 0
    	objDoc.Close
        
  		objVisio.Quit
        Set objVisio = Nothing
              
    	Document.AddRendition "c:\temp\" & PDFName
    
    	'VSD en PDF van tijdelijke lokatie verwijderen  
    	Set objFSO = CreateObject("Scripting.FileSystemObject")
    	If objFSO.FileExists("c:\temp\" & PDFName) = True Then objFSO.DeleteFile "c:\temp\" & PDFName
    	If objFSO.FileExists("c:\temp\" & Document.FileName) = True Then objFSO.DeleteFile "c:\temp\" & Document.FileName
    End If
End Sub

'RAC 21-04-2015 rename of object not allowed
Sub DocGenericEvent_BeforeRename(Batch, NewName)
	If Not Document Is Nothing Then
    	'RAC 21-04-2015
		If ucase(Document.DocumentType) = "OBJECTS" Then
        	If Document.WorkFlowState = AS_WF_RELEASED Then
               	'Batch.Abort "Renaming of object only allowed through TechnicalID property"
                Batch.FailCurrent "Renaming of object only allowed through TechnicalID property"
            End If
        End If
		'Add your code for folder objects
	End If
End Sub

Sub DocCopyMoveEvent_BeforeMove(Batch, TargetFolder)
	'RAC Volantis, Move vanuit master afbreken
	If Not Document Is Nothing Then
    	If client.ImportDetails <> AS_ID_CREATEPROJCOPY And client.ImportDetails <> AS_ID_CREATEPROJCOPYWLOCK  Then
			If Document.Branch = "Master" Then
	            Batch.Abort "Documenten mogen niet verplaatst worden vanuit de Master"
            End If
        End If
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

'********************************************************************************************************************************
'**************************************************************FUNCTIONS*********************************************************
'********************************************************************************************************************************
Function CalculateFileNameGeneralDocument
	If client.ImportDetails <> AS_ID_CREATEPROJCOPY And client.ImportDetails <> AS_ID_RELEASETOMASTER And client.ImportDetails <> AS_ID_CREATEPROJCOPYWLOCK  Then
    	'WinMsgBox "Calc filename General"
		'CalculateFileNameGeneralDocument = 	Left(Document.Reference_Site,3) & "-" & Left(Document.Reference_ProcessCell,4)  & "-" &_
		'							 		Left(Document.Reference_Process,5) & "-" & Document.Reference_DisciplineCode & Document.Floor & "-" &_
		'							 		FormatSequenceNum(Vault.Sequence(Document.Reference_Site & Document.Reference_ProcessCell  & Document.Reference_Process & Document.Reference_DisciplineCode & Document.Floor).Next(1),4) &_
		'							 		FileExtension(Document.FileName)
        If Document.ForceFileName Then
        	If Document.ProcessCell = "PC90 Project" Then
            	'Seq = Vault.Sequence(Document.Projects & Document.Reference_Site & Document.Reference_ProcessCell  & Document.Reference_Process & Document.Reference_DisciplineCode & Document.Floor) + 1
                '2017-06-12, Sequence bepalen op basis van projectanddescription, projects is hier al leeg
                Seq = Vault.Sequence(Document.Project_ProjectAndDescription & Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor) + 1
        		SeqEntered = Cint(Document.DocSequenceNumber)
        		If SeqEntered >= Seq Then
        			'sequence is gelijk of hoger, nieuwe sequence zetten
     				Vault.Sequence(Document.Project_ProjectAndDescription & Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqEntered)      
        		End If
           		CalculateFileNameGeneralDocument = Document.NamePart & Document.DocSequenceNumber & Document.NameExt 
            Else
            	Seq = Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor) + 1
        		SeqEntered = Cint(Document.DocSequenceNumber)
        		If SeqEntered >= Seq Then
        			'sequence is gelijk of hoger, nieuwe sequence zetten
     				Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqEntered)      
        		End If
           		CalculateFileNameGeneralDocument = Document.NamePart & Document.DocSequenceNumber & Document.NameExt 
            End If 
        
        
        	
           	
        Else
        	CalculateFileNameGeneralDocument = Document.FileName 	
        End If
    Else
    	CalculateFileNameGeneralDocument = Document.FileName 
    End If
End Function


Function PreviewFileNameGeneralDocument
		PreviewFileNameGeneralDocument = 	Left(Document.Site,3) & "-" & Left(Document.ProcessCell,4)  & "-" &_
											Left(Document.Process,5) & "-" & Document.DisciplineCode & Document.Floor & "-" & Document.DocSequenceNumber & FileExtension(Document.FileName)
End Function

Function CalculateFileNameObject
	If client.ImportDetails <> AS_ID_CREATEPROJCOPY And client.ImportDetails <> AS_ID_RELEASETOMASTER And client.ImportDetails <> AS_ID_CREATEPROJCOPYWLOCK  Then
    	CalculateFileNameObject = 	Document.Object_TechnicalIDNumber & ".obj"
    Else
    	CalculateFileNameObject = Document.FileName 
    End If
End Function


Function CalculateFileNameControlledDocument
	'WinMsgBox client.ImportDetails
    
    If client.ImportDetails <> AS_ID_CREATEPROJCOPY And client.ImportDetails <> AS_ID_RELEASETOMASTER And client.ImportDetails <> AS_ID_CREATEPROJCOPYWLOCK  Then
	    'WinMsgBox "Calc filename Controlled"
		'CalculateFileNameControlledDocument = 	Left(Document.AsBuilt_Site,3) & "-" & Left(Document.AsBuilt_ProcessCell,4)  & "-" &_
		'							 			Left(Document.AsBuilt_Process,5) & "-" & Document.AsBuilt_DisciplineCode & Document.Floor & "-" &_
		'							 			FormatSequenceNum(Vault.Sequence(Document.AsBuilt_Site & Document.AsBuilt_ProcessCell  & Document.AsBuilt_Process & Document.AsBuilt_DisciplineCode & Document.Floor).Next(1),4) &_
		'							 			FileExtension(Document.FileName)
        
        If Document.ForceFileName Then
           	Seq = Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor) + 1
        	SeqEntered = Cint(Document.DocSequenceNumber)
        
        	If SeqEntered >= Seq Then
        		'sequence is gelijk of hoger, nieuwe sequence zetten
     			Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqEntered)      
        	End If
           	CalculateFileNameControlledDocument = Document.NamePart & Document.DocSequenceNumber & Document.NameExt 
        Else
        	CalculateFileNameControlledDocument = Document.FileName 	
        End If
        
    Else
    	CalculateFileNameControlledDocument = Document.FileName 
    End If
    
End Function

Function PreviewFileNameControlledDocument
		PreviewFileNameControlledDocument = Left(Document.Site,3) & "-" & Left(Document.ProcessCell,4)  & "-" &_
											Left(Document.Process,5) & "-" & Document.DisciplineCode & Document.Floor & "-" 
End Function

Function SetSequence

	If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Or UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then

        Docsplit = split(Document.FileName,"-")
            
		If ubound(Docsplit) = 4 Or ubound(Docsplit) = 5 Then
            	If ubound(Docsplit) = 4 Then
        			Seq =  split(Docsplit(4),".")
        		Else
          			Seq =  split(Docsplit(5),".")
        		End If
            
            	SetSequence = Seq(0)
        		SeqDocument= Cint(SetSequence)
        	
        	If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
        		If SeqDocument > (Vault.Sequence(Document.ite & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).Value) Then
					Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqDocument)
				End If
        	End If
        	If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
        		If SeqDocument > (Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).Value) Then
                	If Document.ProcessCell = "PC90 Project" Then
						Vault.Sequence(Document.Projects & Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqDocument)
                    Else
                    	Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqDocument)
                    End If
            	End If
        	End If
    	End If
	End If
        
End Function
Function SetSequenceNr

	If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Or UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then

        Docsplit = split(Document.FileName,"-")
            
		If ubound(Docsplit) = 4 Or ubound(Docsplit) = 5 Then
            	If ubound(Docsplit) = 4 Then
        			Seq =  split(Docsplit(4),".")
        		Else
          			Seq =  split(Docsplit(5),".")
        		End If
            
            	SetSequenceNr = Seq(0)
        		SeqDocument= Cint(SetSequenceNr)
        	
        	If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
        		'If SeqDocument > (Vault.Sequence(Document.AsBuilt_Site & Document.AsBuilt_ProcessCell  & Document.AsBuilt_Process & Document.AsBuilt_DisciplineCode & Document.Floor).Value) Then
					Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqDocument)
				'End If
        	End If
        	If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
        		'If SeqDocument > (Vault.Sequence(Document.Reference_Site & Document.Reference_ProcessCell  & Document.Reference_Process & Document.Reference_DisciplineCode & Document.Floor).Value) Then
                If Document.ProcessCell = "PC90 Project" Then
					Vault.Sequence(Document.Projects & Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqDocument)
                Else
                	Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(SeqDocument)
            	End If
        	End If
    	End If
	End If
        
End Function

Function SetNamePart
	If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
		SetNamePart = Left(Document.Site,3) & "-" & Left(Document.ProcessCell,4)  & "-" &_
					  Left(Document.Process,5) & "-" & Document.DisciplineCode & Document.Floor & "-" 
	End If
    If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
    	If Document.ProcessCell = "PC90 Project" Then
        	'2017-06-12 RC, projectomschrijving uit bestandsnaam halen
        	Project = Split(Document.Projects," ")
			SetNamePart = Project(0) & "-" & Left(Document.Site,3) & "-" & Left(Document.ProcessCell,4)  & "-" &_
						  Left(Document.Process,5) & "-" & Document.DisciplineCode & Document.Floor & "-" 
        Else
        	SetNamePart = Left(Document.Site,3) & "-" & Left(Document.ProcessCell,4)  & "-" &_
						  Left(Document.Process,5) & "-" & Document.DisciplineCode & Document.Floor & "-" 
        End If
    End If
End Function


Function SetNameExt
	SetNameExt = FileExtension(Document.FileName)
End Function


Function GetSequence
	If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
    	
		GetSequence = FormatSequenceNum((Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).Value + 1),4)
	End If
    If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
    	'WinMsgBox Document.Reference_Site & Document.Reference_ProcessCell  & Document.Reference_Process & Document.Reference_DisciplineCode & Document.Floor
        If Document.ProcessCell = "PC90 Project" Then
        	GetSequence = FormatSequenceNum((Vault.Sequence(Document.Projects & Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).Value + 1),4)
        Else
        	GetSequence = FormatSequenceNum((Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).Value + 1),4)
        End If
    End If
End Function


Function GetUnit
	GetUnit = Vault.Query("Unit").GetValues(Null,Null, "Unit",,"Unit")
End Function

Function GetEquipmentModule
	GetEquipmentModule = Vault.Query("EquipmentModule").GetValues(Null,Null, "EquipmentModule",,"EquipmentModule")
End Function

Function GetProcessCell
	GetProcessCell = Vault.Query("ProcessCell").GetValuesEx("Name","Doctype = '" & Document.DocumentType & "'" , "Name",,"Name")
End Function

Function GetProcess
    If Document.ProcessCell <> "" Then
		GetProcess = Vault.Query ("Process").GetValuesEx("Process","Processcell = '" & Left(Document.ProcessCell,4) & "'" ,,,"Process")
	Else
    	GetProcess = Vault.Query ("Process").GetValuesEx("Process",,,,"Process")
        If Document.ProcessCell = "PC90 Project" Then
        	GetProcess = Vault.Query ("Process").GetValuesEx("Process",,,,"Process")
        End If
	End If
End Function

Function GetDiscipline
	GetDiscipline = Vault.Query("Discipline").GetValues(Null,Null, "Description",,"Description")
End Function

Function GetDisciplineClass
	
	If Document.Discipline  <> "" Then
        GetDisciplineClass = Vault.Query ("DisciplineClass").GetValuesEx("Description","Discipline = '" & Document.Discipline & "'" ,,,"Description")
   	End If
End Function

Function GetDisciplineCode
	If Document.Discipline  <> "" Then
        DiscCode = Vault.Query ("Discipline").GetValuesEx("Code","Description = '" & Document.Discipline & "'" ,,,"Code")
        If IsArray(DiscCode) Then
     	    GetDisciplineCode = DiscCode(0,0)
        End If
    End If
End Function

'Function GetDisciplineCodeMaster
'	If Document.Discipline <> "" Then
'        DiscCode = Vault.Query ("Discipline").GetValuesEx("Code","Description = '" & Document.Discipline & "'" ,,,"Code")
'        If IsArray(DiscCode) Then
'     	    GetDisciplineCodeMaster = DiscCode(0,0)
'        End If
'    End If
'End Function

Function GetBuilding
	GetBuilding = Vault.Query ("Building").GetValues(Null,Null, "Building",,"Building")
End Function

Function GetLocation
	If Document.Object_Building  <> "" Then
        GetLocation = Vault.Query ("Location").GetValuesEx("Location","Building = '" & Document.Object_Building & "'" ,,,"Location")
    End If
End Function

Function GetRoom
	If Document.Object_Location  <> "" Then
        GetRoom = Vault.Query ("Room").GetValuesEx("Room","Location = '" & Document.Object_Location & "'" ,,,"Room")
    End If
End Function

Function GetSection
	If Document.Object_Room  <> "" Then
    	GetSection = Vault.Query ("Section").GetValuesEx("[Section]","Room = '" & Document.Object_Room  & "'" ,,,"[Section]")
    End If
End Function


'Function GetDisciplineClass
'	If Document.Reference_Discipline <> "" Then
'        GetDisciplineClass = Vault.Query ("DisciplineClass").GetValuesEx("Description","Discipline = '" & Document.Reference_Discipline & "'" ,,,"Description")
'    End If
'End Function

Sub ChangeDisciplineCode_Execute(Batch)
	If Not Document Is Nothing Then
		If Document.Discipline <> "" Then
        DiscCode = Vault.Query ("Discipline").GetValuesEx("Code","Description = '" & Document.Discipline & "'" ,,,"Code")
        If IsArray(DiscCode) Then
     	    Document.DisciplineCode = DiscCode(0,0)
        End If
    End If
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Function Archive_Document()
	If Not Document Is Nothing Then
		'document verplaatsen naar de Archive structuur
        Splits = Split(Document.Path,"\")
        i = 1
        Newfolder = "Archive\"
        Do While i < UBound(Splits)
        	Newfolder = NewFolder & Splits(i) & "\"
            i = i + 1
        Loop
        'bij objecten niet de tijd er tussen zetten
        If ucase(Document.DocumentType) <> "OBJECTS" Then
            Newfolder = Newfolder & ModificationTime
        Else
        	NewFolder = Left(NewFolder, Len(NewFolder)-1)
        End If
        Document.ArchiveSet_ModificationDate = ModificationDate
        Document.ArchiveSet_ModificationDateTime_String = ModificationTime
        
        Document.Project_ProjectAndDescription = Document.Project_Projectnummer & " " & Document.Project_Projectomschrijving
        
                
        'winmsgbox Len(Newfolder)
        'Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
        
        Document.ArchiveSet_Archive_Root = Document.Branch
        Document.Branch = "Archive"
        
        'client.goto Document.MasterDocument
        'client.Refresh(AS_RF_CHANGED_CURRENTVIEW)
        'client.Refresh(AS_RF_CHANGED_PROPERTIES)
        
	End If
End Function

Function ModificationDate
	Maand = Month(GMTTime2Local(Document.Modified))
    If Len(Maand) = 1 Then
    	Maand = "0" & Maand
    End If
    Dag = Day(GMTTime2Local(Document.Modified))
    If Len(Dag) = 1 Then
    	Dag = "0" & Dag
    End If
    Uur = Hour(GMTTime2Local(Document.Modified))
    If Len(Uur) = 1 Then
    	Uur= "0" & Uur
    End If
    Minuut = Minute(GMTTime2Local(Document.Modified))    
         
    If Len(Minuut) = 1 Then
    	Minuut = "0" & Minuut
    End If
    ModificationDate = Year(GMTTime2Local(Document.Modified)) & "-" & Maand & "-" & Dag
End Function

Function ModificationTime
	Maand = Month(GMTTime2Local(Document.Modified))
    If Len(Maand) = 1 Then
    	Maand = "0" & Maand
    End If
    Dag = Day(GMTTime2Local(Document.Modified))
    If Len(Dag) = 1 Then
    	Dag = "0" & Dag
    End If
    Uur = Hour(GMTTime2Local(Document.Modified))
    If Len(Uur) = 1 Then
    	Uur= "0" & Uur
    End If
    Minuut = Minute(GMTTime2Local(Document.Modified))    
         
    If Len(Minuut) = 1 Then
    	Minuut = "0" & Minuut
    End If
    ModificationTime = Year(GMTTime2Local(Document.Modified)) & "-" & Maand & "-" & Dag & "_" & Uur & "_" & Minuut 
End Function
'********************************************************************************************************************************
'**************************************************************COMMANDS**********************************************************
'********************************************************************************************************************************
Sub CheckObject_Execute(Batch)
	If Not Document Is Nothing Then
		'Add your code for document objects
        Teller = Vault.FindDocuments(Document.Object_TechnicalIDNumber & ".obj").Count
        
        If ( Teller >= 1) Then
        	WinMsgBox "Object is " & Teller & " times found, object is NOT unique"
        Else
        	WinMsgBox "Ok, object name is unique"
        End If
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Sub CheckDocument_Execute(Batch)
	If Not Document Is Nothing Then
		Teller = Vault.FindDocuments(Document.NamePart & Document.DocSequenceNumber & Document.NameExt).Count
        
        If ( Teller >= 1) Then
        	WinMsgBox "Document is " & Teller & " times found, document is NOT unique"
        Else
        	WinMsgBox "Ok, document name is unique"
        End If   
	End If
End Sub

Sub Folder_Execute(Batch)
	If Not Document Is Nothing Then
		'Add your code for document objects
        winmsgbox Document.ParentProject.Name 
        
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Sub Folder_AfterWizard(Batch)
	If Not Document Is Nothing Then
		'Add your code for document objects
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Sub ArchiveDocument_Execute(Batch)
	If Not Document Is Nothing Then
		'document verplaatsen naar de Archive structuur
        'winmsgbox HuidigeDatumTijd 'Datum & "_" & Hour(Today) & "_" & Minute(Today) & "_" & Second(Today)
        'winmsgbox Document.Path
        Splits = Split(Document.Path,"\")
        i = 1
        Newfolder = "Archive\"
        Do While i < UBound(Splits)
        	Newfolder = NewFolder & Splits(i) & "\"
            i = i + 1
        Loop
        Newfolder = Newfolder & ModificationTime
        
        'winmsgbox Len(Newfolder)
        Document.MoveTo vault.RootFolder.GetSubFolder(NewFolder ,,AS_NFO_CREATE_IFNOTEXIST)
        
        Document.ChangeWorkflowState AS_WF_RETIRED, "", ""
        
        
        ElseIf Not Folder Is Nothing Then
    
	End If
End Sub
'********************************************************************************************************************************
'********************************SAVE PROPERTIES*********************************************************************************
'********************************************************************************************************************************
Sub StoreDocumentPropertyForThisBatch(Batch, Document, DocType)
    ' The 'Argument' property of the 'Batch' object is used to store information within the context 
    ' of the current batch operation.

    ' Store the values entered for this document to reuse them for the next document(s) in the batch
    'Batch.Argument ("ProductCategory") = Document.ProductCategory
    Batch.Argument ("Site") = Document.Site
    'WinMsgBox "Store " & Document.Reference_ProcessCell
    Batch.Argument ("Reference_Site") = Document.Reference_Site
    Batch.Argument ("ProcessCell") = Document.ProcessCell
    Batch.Argument ("Reference_ProcessCell") = Document.Reference_ProcessCell
    Batch.Argument ("Process") = Document.Process
    Batch.Argument ("Reference_Process") = Document.Reference_Process
    Batch.Argument ("Discipline") = Document.Discipline
    Batch.Argument ("Reference_Discipline") = Document.Reference_Discipline
    Batch.Argument ("Reference_DisciplineCode") = Document.Reference_DisciplineCode
    Batch.Argument ("DisciplineCode") = Document.DisciplineCode
    Batch.Argument ("DisciplineClass") = Document.DisciplineClass
    Batch.Argument ("Reference_DisciplineClass") = Document.Reference_DisciplineClass
    Batch.Argument ("TitleBlockLine1") = Document.TitleBlockLine1
    Batch.Argument ("TitleBlockLine2") = Document.TitleBlockLine2
    Batch.Argument ("TitleBlockLine3") = Document.TitleBlockLine3
        
    ' Some properties might depend on the document type
    ' So we store the value in a argument that includes the document type name
    'Batch.Argument ("Material_" & DocType.InternalName) = Document.Material 
    
End Sub

' Custom procedure used in the 'BeforeNewDocument' event
Sub InitializeDocumentProperty(Batch, Document, DocType)
    ' Initialize the new document with the values stored from a previous document
    Document.Site = Batch.Argument ("Site") 
    If Batch.IsFirstInBatch And Document.Reference_Site = "" Then
	    Document.Reference_Site = Batch.Argument ("Reference_Site") 
    End If
    'Document.ProcessCell = Batch.Argument ("ProcessCell") 
    If Batch.IsFirstInBatch And Document.ProcessCell = "" Then
		Document.ProcessCell = Batch.Argument ("ProcessCell")  
    End If
    'Document.Process = Batch.Argument ("Process") 
    If Batch.IsFirstInBatch And Document.Process = "" Then
    	Document.Process = Batch.Argument ("Process") 
    End If
    'Document.Discipline = Batch.Argument ("Discipline") 
    If Batch.IsFirstInBatch And Document.Discipline = "" Then
    	Document.Discipline = Batch.Argument ("Discipline") 
   	End If
    Document.Reference_DisciplineCode = Batch.Argument ("Reference_DisciplineCode") 
    Document.DisciplineCode = Batch.Argument ("DisciplineCode") 
    'Document.DisciplineClass = Batch.Argument ("DisciplineClass") 
	If Batch.IsFirstInBatch And Document.DisciplineClass = "" Then
    	Document.DisciplineClass = Batch.Argument ("DisciplineClass") 
    End If
    Document.TitleBlockLine1 = Batch.Argument ("TitleBlockLine1") 
    Document.TitleBlockLine2 = Batch.Argument ("TitleBlockLine2") 
    Document.TitleBlockLine3 = Batch.Argument ("TitleBlockLine3") 
End Sub

' Custom function used in the 'BeforeNewDocument' and 'AfterNewDocument' event
Function DoYouWantToInitializeThePropertyValues(DocType)
    DoYouWantToInitializeThePropertyValues = True
    'If DocType.InternalName = "GenericDocument" Then 
    '	' For example exclude Generic Documents
    '    DoYouWantToInitializeThePropertyValues = False
    'End If
End Function


Function GetSubFoldersProject
	NewFolder = "Project\NIJ"
	'NewFolder = Document.ParentFolder.Path & "\Master\" & Document.ProcessCell & "\" & Document.Process & "\" & Document.Discipline & "\" & Document.DisciplineClass
    GetSubFoldersProject vault.RootFolder.GetSubFolder("Project\NIJ",,AS_NFO_NOFOLDER_ERROR ).GetSubFolderNames 
End Function
Function GetProjectFolders(strProjectRoot)
		GetProjectFolders = vault.RootFolder.GetSubFolder("Project\NIJ", , AS_NFO_NOFOLDER_ERROR).GetSubFolderNames()
       
End Function

Sub TaggedTrue_Execute(Batch)
	If Not Document Is Nothing Then
		'Add your code for document objects
        Document.Object_Tagged = True
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Sub DocGenericEvent_BeforeDelete(Batch)
	If Not Document Is Nothing Then
		'controleren of document niet in Master / Reference / Object staat
        If Document.Branch <> "Project" And Document.Branch <> "Archive" Then
        	If ucase(vault.User) <> "RAC" Then
        		WinMsgBox "It is not allowed to delete a document, use the Retire function to Archive the document"
            	Batch.Abort 
            End If
        End If
    ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub
Sub CheckDoc_Execute(Batch)
	If Not Document Is Nothing Then
		'Add your code for document objects
        'winmsgbox Document.ID 
		'Set cDoc = vault.GetDocument("{ef6d1793-6796-11e1-0000-27c2899aec96}")
         Set cDoc = vault.GetDocument("{ef6d1795-6796-11e1-0000-27c2899aec96}")
        'If cDoc Is Not Nothing Then
        	WinMsgBox cDoc.FileName
        'End If
        
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub


Sub DocWorkingCopyEvent_BeforeCreateWC(Batch)
	'Add your code here
     If Document.Branch = "Master" Then
	    WinMsgBox "It is not allowed to edit documents in the Master, please create a copy to a project"
        Batch.Abort 
    End If
End Sub


Sub ShowOriginalFilename_Execute(Batch)
	If Not Document Is Nothing Then
		'Add your code for document objects
        WinMsgBox Document.Property("AMFSObjectPropertySet._IMPORTEDFROM")
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Function GetProjects
	'GetProjects = Vault.Table ("Projects").GetValues(Null,Null, "ProjectnumberDescription",,"ProjectnumberDescription")
    
    projfolders = vault.RootFolder.GetSubFolder("Project\NIJ",,AS_NFO_NOFOLDER_ERROR ).GetSubFolderNames
    
    Teller = 0
    Dim Projects()
   
    For Each fld in projfolders
    	ReDim PRESERVE Projects(Teller)
        Set tempFLD = vault.RootFolder.GetSubFolder("Project\NIJ\" & fld)
        Projects(Teller) = fld & " " & tempFLD.Project_Projectomschrijving      
        Teller = Teller + 1
    Next
    GetProjects = Projects
    
End Function


Sub CreateRoomsintable_Execute(Batch)
	If Not Document Is Nothing Then
		
     
            Room = Document.Object_TechnicalIDNumber & " " & Document.Object_TechnicalIDDescription
            'WinMsgBox Room
            'kijken of room al in tabel bestaat
            CheckRoom = Vault.Query("Room").GetValuesEx("Room","Room = '" & Room  & "'" ,,,"Room")
            If Not Isarray(CheckRoom) Then
            	'Room toevoegen
                'WinMsgBox "Room toevoegen: " & Document.Object_Location & " " & Room
                Vault.Query("Room").AddValues Array("Location","Room","OldRoom"),Array(Document.Object_Location,Room,"New")
            End If
           
     
    
	End If
End Sub

Sub OldnameinTitle1_Execute(Batch)
	If Not Document Is Nothing Then
    	Title1 = Document.Property("AMFSObjectPropertySet._IMPORTEDFROM")
        
        Point = InStrRev(Title1 ,".",-1, vbTextCompare)
    
    	Title1 = Left(Title1 ,(Point-1))
    
		Document.TitleBlockLine1 = Title1 
        
        
        
        
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Sub SetSequence_Execute(Batch)
	If Not Document Is Nothing Then
		If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Or UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
          	If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
        		'If SeqDocument > (Vault.Sequence(Document.Reference_Site & Document.Reference_ProcessCell  & Document.Reference_Process & Document.Reference_DisciplineCode & Document.Floor).Value) Then
                If Document.ProcessCell = "PC90 Project" Then
					Vault.Sequence(Document.Projects & Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(Document.DocSequenceNumber)
                Else
                	Vault.Sequence(Document.Site & Document.ProcessCell  & Document.Process & Document.DisciplineCode & Document.Floor).SetTo(Document.DocSequenceNumber)
            	End If
        	End If
    	End If
	End If
End Sub


Sub UpdateTitle_Execute(Batch)
	If Not Document Is Nothing Then
	    If (Document.Title <> Trim(Document.TitleBlockLine1 & " " & Document.TitleBlockLine2 & " " & Document.TitleBlockLine3 & " " & Document.TitleBlockLine4))  Then
          	Document.Title = Trim(Document.TitleBlockLine1 & " " & Document.TitleBlockLine2 & " " & Document.TitleBlockLine3 & " " & Document.TitleBlockLine4)
        End If
        
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Sub UpdateRoomName_Execute(Batch)
	If Not Document Is Nothing Then
		'OXN Roomname bijwerken naar nieuwe benaming
       	NewRoomValue = Vault.Query("Room").GetValues(Array("Location", "OldRoom"), Array(Document.Object_Location, Document.Object_Room), Array("Room"))
       	'WinMsgBox(NewRoomValue(0,0))
       	Document.Object_Room = NewRoomValue(0,0)
	End If
End Sub

Sub UpdateProjectnummer_Execute(Batch)
	If Not Document Is Nothing Then
		nummer = WinInputBox("Nummer")
        oms = WinInputBox("Omschrijving")
        
        Document.Project_ProjectAndDescription = nummer & " " & oms
        Document.Project_Projectomschrijving = oms
	End If
End Sub

Sub SetProperties_Execute(Batch)
If Not Document Is Nothing Then

		'If  Document.Rendition <> "" Then
        '	Document.TitleBlockLine4 = "new"
        'End If
        'If Document.Project_Project_Root = "Master" Then
        '	Document.TitleBlockLine4 = "new"
        'End If
        
		'If Document.TitleBlockLine4 = "new" Then
        '	Document.TitleBlockLine4 = vbNullString
        'End If    

       'If Document.Branch = "Archive" Then
        '	Dim pad
        '	pad = split(Document.path,"\")
       ' 	Document.ArchiveSet_Archive_Root = pad(2)
        '	Document.ArchiveSet_ModificationDateTime_String = ModificationTime
       ' End If
       ' Document.Site = "NIJ"
        If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
             Document.ProcessCell = Document.Process
             Document.Process = Document.Discipline
              Document.Discipline = Document.AsBuilt_Discipline
             Document.DisciplineClass = Document.AsBuilt_DisciplineClass
             Document.DisciplineAndClass = Document.AsBuilt_Discipline & " - " & Document.AsBuilt_DisciplineClass
        End If
        'If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
 	'		  Document.ProcessCell = Document.Reference_ProcessCell
    '         Document.Process = Document.Reference_Process
    '         Document.Discipline = Document.Reference_Discipline
    '         Document.DisciplineClass = Document.Reference_DisciplineClass
    '         Document.DisciplineAndClass = Document.Reference_Discipline & " - " & Document.Reference_DisciplineClass
    '    End If
    '    If ucase(Document.DocumentType) = "OBJECTS" Then
   ' 		  Document.ProcessCell = Document.Object_ProcessCell
   '           Document.Process = Document.Object_Process
   '       
   '     End If
       
       
    '   If UCase(Document.DocumentType) = "CONTROLLED DOCUMENTS" Then
     '  	Document.DisciplineCode = Getdisciplinecode
     '  End If
     '  If UCase(Document.DocumentType) = "GENERAL DOCUMENTS" Then
     '   Document.DisciplineCode = Getdisciplinecode
     '  End If
        
	ElseIf Not Folder Is Nothing Then
    
    	'Folder.Project_ProjectAndDescription = Folder.Project_Project & " " & Folder.Project_Projectomschrijving
    
    	'Folder.Project_Projectnummer = Folder.Name
        'Folder.Project_Project = Folder.Name
        'WinMsgBox ValidateFolderName(Folder.Project_Project & " " & Folder.Project_Projectomschrijving,"-")
        'Folder.Property("Project.Projectomschrijving")
        'WinMsgBox Folder.Project_Projectomschrijving
         'Folder.Name = ValidateFolderName(Folder.Project_Project & " " & Folder.Project_Projectomschrijving,"-")
        'WinMsgBox Folder.Property("Project.Projectomschrijving")
        'Call Vault.Table("Projects").AddValues(Array("Projectnumber", "Projectdescription", "ProjectnumberDescription"), Array(Folder.Project_Project,Folder.TitleBlockLine1,Folder.Project_Project & " " & Folder.TitleBlockLine1))
        
        'Folder.Project_Projectomschrijving = Folder.TitleBlockLine1
        
	End If

	
	
	End Sub


Sub FolderGenericEvent_AfterNewFolder(Batch, Action)
	'Add your code here
    If Folder.IsProject Then
    	Folder.TitleBlockLine1 = Folder.Project_Projectomschrijving
        Folder.Project_ProjectAndDescription = Folder.Project_Project & " " & Folder.Project_Projectomschrijving
    End If
End Sub

Function RenditionPage_IsVisible()
	If Not Document Is Nothing Then
		RenditionPage_IsVisible = True
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Function

Sub Objectenkoppelen_Execute(Batch)
	If Not Document Is Nothing Then
    	Dim Objects
        Dim Object
    	' Get clipboard text
		Set objHTML = CreateObject("htmlfile")
		Objects = objHTML.ParentWindow.ClipboardData.GetData("text")
        
        'Objects = WinInputBox("Geef de objecten op die je wil koppelen")
        
        If Batch.IsFirstInBatch Then
        	WinMsgBox "Objecten gelezen uit geheugen: " & vbCrLf & Objects
        	Answer = WinMsgBox( "Verder gaan om te koppelen" , AS_YesNo , "Meridian" )
        	Batch.Argument ("UserAnswer") = Answer 
    		If Batch.Argument ("UserAnswer") = AS_No Then
            	Exit Sub
            End If
        	Batch.Argument("Gevonden") = "Gekoppelde objecten:"
            Batch.Argument("NietGevonden") = "Niet gevonden objecten:"
        End If 
        
        If Len(Objects) > 0 Then
        	ObjectList = Split(Objects,vbCrLf)
            
            For Each Object in ObjectList             
                            
              Object = Object & ".obj"        
             If Vault.FindDocuments(Object,"Objects",,False).Count = 1 Then
        		'Object gevonden, als er nog geen koppeling is deze koppelen
        		For Each Obj In Vault.FindDocuments(Object,"Objects",,False)
        			'Object koppelen 
                	If Not Document.GetReferences("TagObjectReference",False).Exist(Obj.ID) Then
        				Document.GetReferences("TagObjectReference",False).Add(Obj.ID)
                	End If
        	 	Next 
                'log gekoppeld
                Batch.Argument("Gevonden") = Batch.Argument("Gevonden") & " " & Object
             Else
            	'Object niet gevonden     	
                'log als niet gevonden
				If Trim(Object) <> ".obj" Then
	                Batch.Argument("NietGevonden") = Batch.Argument("NietGevonden") & " " & Object
                End If
             End If
            Next
        Else
        	WinMsgBox "Geen invoer, functie wordt afgebroken" 
        	Exit Sub
        End If
        
        If Batch.IsLastinBatch Then
        	If Trim(Batch.Argument("Gevonden")) <> "Gekoppelde objecten:" Then
        		Batch.PrintDetails Batch.Argument("Gevonden")
            End If
            If Trim(Batch.Argument("NietGevonden")) <> "Niet gevonden objecten:" Then
	            Batch.PrintDetails Batch.Argument("NietGevonden")
            End If
        End If
   
	End If
End Sub

Sub VisioPDFbijwerken_Execute(Batch)
	If Not Document Is Nothing Then
	'Als een VSD document released wordt dan de PDF als rendition opslaan
    If Right(Ucase(Document.FileName),3) = "VSD" Then
    	Document.SaveToFile "c:\temp\" & Document.FileName 
             
    	Set objVisio = CreateObject("Visio.Application")
    	Set objDoc = objVisio.Documents.OpenEx("c:\temp\" & Document.FileName, &H2 + &H10 + &H40)    
    
    	PDFName = Document.FileName
    	PDFName = Left(PDFName,Len(PDFName)-3) & "pdf"
     
    	objDoc.ExportAsFixedFormat 1, "c:\temp\" & PDFName, 1, 0
    	objDoc.Close
        
  		objVisio.Quit
        Set objVisio = Nothing
              
    	Document.AddRendition "c:\temp\" & PDFName
    
    	'VSD en PDF van tijdelijke lokatie verwijderen  
    	Set objFSO = CreateObject("Scripting.FileSystemObject")
    	If objFSO.FileExists("c:\temp\" & PDFName) = True Then objFSO.DeleteFile "c:\temp\" & PDFName
    	If objFSO.FileExists("c:\temp\" & Document.FileName) = True Then objFSO.DeleteFile "c:\temp\" & Document.FileName
    End If
	End If
End Sub


Sub ChangeRoomdata_Execute(Batch)
	If Not Document Is Nothing Then
        '2-16-2017 OXN: Roomname bijwerken naar nieuwe benaming
       Document.tmpRoom = Document.tmpTechnicalID & " " & Document.tmpTechnicalDescription      

        If Document.tmpBuilding <> Document.Object_Building Then
            WinMSgBox("Building Changed")      
            'Kluis doorlopen en Building aanpassen waar gebruikt
            Criteria = Array(Array("Object.Room",  IC_OP_EQUALS, Document.Object_TechnicalIDNumberANDDescription))
            For Each Obj In Vault.FindDocuments(,"Objects",Criteria,False) 
                Obj.Object_Building = Document.tmpBuilding
                Obj.ApplyPropertyValues
            Next
            WinMsgBox("Objects have been updated")
        End If

        If Document.tmpLocation <> Document.Object_Location Then
            WinMSgBox("Location Changed")
            'Kluis doorlopen en Location aanpassen waar gebruikt
            Criteria = Array(Array("Object.Room",  IC_OP_EQUALS, Document.Object_TechnicalIDNumberANDDescription))
            For Each Obj In Vault.FindDocuments(,"Objects",Criteria,False) 
                Obj.Object_Location = Document.tmpLocation
                Obj.ApplyPropertyValues
            Next
            WinMsgBox("Objects have been updated")          
        End If 
         
        If Document.tmpRoom <> Document.Object_TechnicalIDNumberANDDescription Then
            WinMSgBox("Room description changed")   
            'Kluis doorlopen en Room aanpassen waar gebruikt
            Criteria = Array(Array("Object.Room",  IC_OP_EQUALS, Document.Object_TechnicalIDNumberANDDescription))
            For Each Obj In Vault.FindDocuments(,"Objects",Criteria,False)          
                Obj.Object_Room = Document.tmpRoom
                Obj.ApplyPropertyValues
            Next 
            WinMsgBox("Objects have been updated")
        End If
        
        'Table aanpassen
        Vault.Query("Room").UpdateValues Array("Location", "Room", "OldRoom"), Array(Document.Object_Location, Document.Object_TechnicalIDNumberANDDescription, "New"), Array("Location", "Room", "OldRoom"), Array(Document.tmpLocation, Document.tmpRoom, "New")

        'Afhandelen Location Object
        Document.Title = Document.tmpTechnicalDescription
        Document.TitleBlockLine1 = Document.tmpTechnicalDescription
        Document.Object_TechnicalIDDescription = Document.tmpTechnicalDescription
        Document.Object_TechnicalIDNumberANDDescription = Document.tmpRoom 
        Document.Object_Location = Document.tmpLocation
        Document.Object_Building = Document.tmpBuilding

        'Tijdelijke waardes opruimen
        Document.tmpBuilding = Empty
        Document.tmpLocation = Empty
        Document.tmpRoom = Empty
        Document.tmpTechnicalDescription = Empty
        Document.tmpTechnicalID = Empty

	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Function ChangeRoomdata_State(Mode)
    ChangeRoomdata_State = AS_CMD_HIDDEN
    If Not Document Is Nothing Then
        'Add your code for document objects
        If Document.Process = "00007 Building" Then
            ChangeRoomdata_State = AS_CMD_NORMAL
        End If  
    ElseIf Not Folder Is Nothing Then
        'Add your code for folder objects
    End If
End Function

Sub VerifyRoomdata_Execute(Batch)
	If Not Document Is Nothing Then
    
		'Add your code for document objects
         chkroom = Vault.Query("Room").GetValues(Array("Location", "Room"), Array(Document.Object_Location, Document.Object_TechnicalIDNumberANDDescription), Array("Room"))
         If Not IsArray(chkroom) Then
 			WinMSgBox(Document.Object_TechnicalIDNumberANDDescription)
         End If
	ElseIf Not Folder Is Nothing Then
		'Add your code for folder objects
	End If
End Sub

Function VerifyRoomdata_State(Mode)
    VerifyRoomdata_State = AS_CMD_HIDDEN
    If Not Document Is Nothing Then
        'Add your code for document objects
        If Document.Process = "00007 Building" Then
            VerifyRoomdata_State = AS_CMD_NORMAL
        End If  
    ElseIf Not Folder Is Nothing Then
        'Add your code for folder objects
    End If
End Function

Sub SetRenderProps_Execute(Batch)
	If Not Document Is Nothing Then
		Document.Property("BCREnditionPropertySet._PAGELAYOUT") = Document.RenderTemp_PageLayout
        Document.Property("BCREnditionPropertySet._PAGESIZE") = Document.RenderTemp_PageSize
        Document.Property("BCREnditionPropertySet._PAGEORIENTATION") = Document.RenderTemp_PageOrientation
        Document.Property("BCREnditionPropertySet._RENDERCOLOR") = Document.RenderTemp_RenderColor
        Document.Property("BCREnditionPropertySet._PENSETTINGSFILE") = Document.RenderTemp_PenSettings
        Document.Property("BCREnditionPropertySet._XOD_RESOLUTION") = Document.RenderTemp_XOD
        Document.Property("BCREnditionPropertySet._PRINTING_QUALITY") = Document.RenderTemp_PrintQuality
        Document.Property("BCREnditionPropertySet._LAYERTRANSTABLE") = Document.RenderTemp_LayerTranslationTable
        Document.ApplyPropertyValues
	End If
End Sub

Sub DocWorkingCopyEvent_AfterSubmitWC(Batch)
	Document.UpdateRendition
End Sub

Function AIMS_AddComment(commentText, attachmentType, numberOfComments)
	Document.Has_Comments = True
End Function

Function AIMS_DeleteComment(commentText, attachmentType, numberOfComments)
Document.Log "Comment deleted by: " & User.Name
	If numberOfComments = 0 Then
	Document.Has_Comments = False
	End If
End Function

