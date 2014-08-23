
On Error resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const TristateUseDefault = 2
Const TristateTrue = 1
Const TristateFalse = 0

Class ini_file
    Dim objSectionDictionary 
    Dim objFileSystemObject
    Dim objRegex

        Private Sub Class_Initialize() 
            Set objSectionDictionary  = CreateObject("scripting.dictionary")
            Set objFileSystemObject = CreateObject("Scripting.Filesystemobject")
            Set objRegex = New RegExp
            objRegex.Global = False
        End Sub

        Private Sub Class_Terminate() 
            Set objSectionDictionary  = Nothing
            Set objFileSystemObject = Nothing
            Set objRegex = Nothing
        End Sub

        Public Function GetSetting(strSection,strKey)
            Set Section = objSectionDictionary (strSection)
            keysForSection = Section.keys
 			
 			If Err.Number <> 0 Then 
 				Err.Clear
 				GetSetting = ""
 			
 			Else
 				GetSetting = Section.Item(strKey)
 		    
 		    End If
 		    
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
            Dim CurrentSection
        
            Do Until objFile.AtEndofStream
                FileAsString = objFile.ReadLine()
                objRegex.Pattern = "^(?!=)\[.*\]"

                If objRegex.Test(FileAsString) Then
                    Debug.WriteLine FileAsString
                    
                    If objSectionDictionary.Exists(replace(replace(FileAsString,"[",""),"]","")) = 0 then
                        objSectionDictionary.Add replace(replace(FileAsString,"[",""),"]",""),CreateObject("scripting.dictionary")
                        CurrentSection = replace(replace(FileAsString,"[",""),"]","")
                        Debug.WriteLine FileAsString
                    End if
                Else
                    objRegex.Pattern = "^(?!;).*?="
                    Set colMatches = objRegex.execute(FileAsString)
                        
                    For Each match In colMatches
                        If objSectionDictionary.Exists(CurrentSection) Then
                            Set settingsForSection = objSectionDictionary.Item(CurrentSection)
                            
                            If settingsForSection.Exists(Replace(match.value,"=","")) = 0 Then
                                settingsForSection.Add Replace(match.value,"=",""),objRegex.replace(FileAsString,"")
                            End If
                        End If
                    Next
                End If
            Loop
        End Function 

End Class

