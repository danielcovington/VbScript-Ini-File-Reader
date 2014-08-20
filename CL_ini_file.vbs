Dim ProgramSettings 
Set ProgramSettings = New ini_file
ProgramSettings.OpenIniFile("C:\Windows\Prefetch\Layout.ini")
MsgBox ProgramSettings.GetSetting("PERF_.NET CLR Networking","First Counter")
Set ProgramSettings = Nothing


Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const TristateUseDefault = 2
Const TristateTrue = 1
Const TristateFalse = 0

Class ini_file


	Dim objSettingsDictionary
	Dim objFileSystemObject
	Dim objIniTextDFile
	Dim objRegex



		Private Sub Class_Initialize() 

			Set objSettingsDictionary = CreateObject("scripting.dictionary")
      		Set objFileSystemObject = CreateObject("Scripting.Filesystemobject")
      		Set objRegex = New RegExp
      		objRegex.Global = False
	
		End Sub

		Private Sub Class_Terminate() 
    
		Set objSettingsDictionary = Nothing
		Set objFileSystemObject = Nothing
		Set objRegex = Nothing
		
		End Sub

		Public Function GetSetting(strSection,strKey)
 	
 			Set Section = objSettingsDictionary(strSection)
 			duh = Section.keys
 			MsgBox UBound(duh)
 			For i = 0 To UBound(duh)
 			 Debug.WriteLine duh(i)
 			 Next
 			 
 			GetSetting = Section.Item(strKey)
 		
		End Function
	
		Public Function OpenIniFile(strFilePath)
		
			If objFileSystemObject.FileExists(strFilePath) Then
			
				Set objIniTextFile = objFileSystemObject.OpenTextFile(strFilePath,ForReading,False,TriStateDefault)
				ParseSections(objIniTextFile)
				OpenIniFile = 1
		
			Else 
		
				OpenIniFile = 0 	
			
			End If
		
		
		End Function
	
		Private Function ParseSections(objFile)
		
			Dim FileAsString
			Dim CurrentSetting
		
		
				Do Until objFile.AtEndofStream
		
					FileAsString = objFile.ReadLine()
					objRegex.Pattern = "^(?!=)\[.*\]"
	
						If objRegex.Test(FileAsString) Then
							Debug.WriteLine FileAsString
							If objSettingsDictionary.Exists(replace(replace(FileAsString,"[",""),"]","")) = 0 then
							objSettingsDictionary.Add replace(replace(FileAsString,"[",""),"]",""),CreateObject("scripting.dictionary")
							CurrentSetting = replace(replace(FileAsString,"[",""),"]","")
							Debug.WriteLine FileAsString
							End if
						Else
							
							objRegex.Pattern = "^(?!;).*?="
							Set colMatches = objRegex.execute(FileAsString)
								
								For Each match In colMatches
		
									If objSettingsDictionary.Exists(CurrentSetting) Then
										Set UGH = objSettingsDictionary.Item(CurrentSetting)
										If UGH.Exists(Replace(match.value,"=","")) = 0 Then
										'Debug.WriteLine Replace(match.value,"=","")
										'Debug.WriteLine " " & objRegex.replace(FileAsString,"")
										UGH.Add Replace(match.value,"=",""),objRegex.replace(FileAsString,"")
										End IF
									End If
								
								Next
						
						End If
		
		
				Loop
		
		End Function 

End Class
