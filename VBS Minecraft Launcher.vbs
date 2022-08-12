Const LauncherName = "VBS Minecraft Launcher"
Const FilesBasePath = "http://fieme.net/MCClassic/"

Function CreateLauncherStructure()
	Set fileSystem = CreateObject("Scripting.FileSystemObject")
	If fileSystem.FolderExists("MCClassic") = False Then
		fileSystem.CreateFolder("MCClassic")
	End If
	
	If fileSystem.FolderExists("MCClassic/Libs") = False Then
		fileSystem.CreateFolder("MCClassic/Libs")
	End If
	
	If fileSystem.FolderExists("MCClassic/Libs/Natives") = False Then
		fileSystem.CreateFolder("MCClassic/Libs/Natives")
	End If
End Function

' "Borrowed" from https://stackoverflow.com/questions/204759/http-get-in-vbs
Function PerformWebDownload(url, output)
	Set webRequest = CreateObject("MSXML2.XMLHTTP.3.0")
	Set webRequestStream = CreateObject("ADODB.Stream")
	call MsgBox("Downloading from """ & url & """ into """ & output & """...", vbOKOnly + vbInformation, LauncherName & " - Download")
	
	On Error Resume Next
	call webRequest.Open("GET", url, False)
	call webRequest.Send()
	
	If Err.Number <> 0 Then
		ResponseStatusCode = 0
	Else
		ResponseStatusCode = webRequest.Status
	End If

	If ResponseStatusCode = 200 Then
		call webRequestStream.Open()
		webRequestStream.Type = 1
		webRequestStream.Write = webRequest.responseBody
		call webRequestStream.SaveToFile(output)
		call webRequestStream.Close()
		' Good response
		PerformWebDownload = 0
	Else
		call MsgBox("An error has occured whilst downloading from """ & url & """! (Status Code: " & ResponseStatusCode & ")", vbOKOnly + vbCritical, LauncherName & " - Download Error")
		' Error response
		PerformWebDownload = 1
	End If
End Function

Function DownloadFiles()
	downloadDialog = MsgBox("Would you like to download the required files for Minecraft Classic?",_
		vbYesNo + vbQuestion, LauncherName & " - Download files")

	If downloadDialog = vbYes Then
		call PerformWebDownload(FilesBasePath & "minecraft.jar", "MCClassic/minecraft.jar")
		call PerformWebDownload(FilesBasePath & "Libs/lwjgl.jar", "MCClassic/Libs/lwjgl.jar")
		call PerformWebDownload(FilesBasePath & "Libs/lwjgl.jar", "MCClassic/Libs/lwjgl_util.jar")
		call PerformWebDownload(FilesBasePath & "Libs/lwjgl.jar", "MCClassic/Libs/jinput.jar")
		call PerformWebDownload(FilesBasePath & "Libs/Natives/OpenAL32.dll", "MCClassic/Libs/Natives/OpenAL32.dll")
		call PerformWebDownload(FilesBasePath & "Libs/Natives/OpenAL64.dll", "MCClassic/Libs/Natives/OpenAL64.dll")
		call PerformWebDownload(FilesBasePath & "Libs/Natives/jinput-dx8.dll", "MCClassic/Libs/Natives/jinput-dx8.dll")
		call PerformWebDownload(FilesBasePath & "Libs/Natives/jinput-dx8_64.dll", "MCClassic/Libs/Natives/jinput-dx8_64.dll")
		call PerformWebDownload(FilesBasePath & "Libs/Natives/jinput-raw.dll", "MCClassic/Libs/Natives/jinput-raw.dll")
		call PerformWebDownload(FilesBasePath & "Libs/Natives/jinput-raw_64.dll", "MCClassic/Libs/Natives/jinput-raw_64.dll")
		call PerformWebDownload(FilesBasePath & "Libs/Natives/lwjgl.dll", "MCClassic/Libs/Natives/lwjgl.dll")
		call PerformWebDownload(FilesBasePath & "Libs/Natives/lwjgl64.dll", "MCClassic/Libs/Natives/lwjgl64.dll")
	End If
End Function

Function ContinueExecution()
	continueExecutionDialog = MsgBox("Would you like to start Minecraft Classic?",_
		vbYesNo + vbQuestion, LauncherName & " - Start Minecraft")
		
	If continueExecutionDialog = vbYes Then
		ContinueExecution = 0
	Else
		ContinueExecution = 1
	End If
End Function

Function StartMinecraft()
	JavaEXEPath = InputBox("Please enter the path to the Java JRE.", LauncherName & " - Select your JRE",_
		"C:\Program Files\Common Files\Oracle\Java\javapath\java.exe")
	Set shell = CreateObject("Wscript.Shell")
	mcRunCMD = """" & JavaEXEPath & """" & " -Djava.library.path=MCClassic/Libs/Natives/ " &_ 
		"-cp MCClassic\Libs\jinput.jar;MCClassic\Libs\lwjgl.jar;MCClassic\Libs\lwjgl_util.jar;MCClassic\minecraft.jar" &_
		" com.mojang.minecraft.Minecraft"
	'MsgBox(mcRunCMD)
	call shell.Run(mcRunCMD)
End Function

call CreateLauncherStructure()
call DownloadFiles()

If ContinueExecution() = 0 Then
	call StartMinecraft()
End If