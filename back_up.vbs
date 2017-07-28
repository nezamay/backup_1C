Option Explicit

Dim fso	'������ ��� ������ � �������� �������� - Scripting.FileSystemObject
DIm Shell '������ WScript.Shell
Dim ConfigFilePath '���� � ����������������� �����
Dim ConfigXML '����������� � ����������� (��� ������� �����) ����. ���� - MSXML2.DOMDocument
Dim CurrentScriptFolder '����������, ��� ���������� ������
Dim Debug '����� ������ � ��� ��������� ���������
Dim TempUploadFolder '��������� ���������� ��� ��������� � ��� ��������
Dim Core1CPath '���� � ������������ ����� ���� 1�
Dim ComConnector8 'COM-���������� � 1�8.2 - V82.COMConnector
Dim Ftp '��������� ������ ��� ������ � ftp - ChilkatFTP.ChilkatFTP
'------------------------------
Class FtpClass
	Private compFtp
	Private state '��� ���������, true/false
	Private data  '��� ������������ ������, �����. �������
	'-----
	Public Property Set Component(value)
		Set compFtp = value
	End Property
	'-----
	Private Sub Class_Initialize()
		InitState null
	End Sub
	'-----
	Private Sub Class_Terminate()
		Disconnect
	End Sub
	'-----
	Public Function GetState(name)
		If state.Exists(name) Then
			GetState = state.Item(name)
		Else
			GetState = false
		End If
	End Function
	'-----
	Private Function LetState(name,value)
		If state.Exists(name) Then
			state.Item(name) = value
		Else
			state.Add name, value
		End If
	End Function
	'-----
	Private Function InitState(name)
		If IsNull(name) Then
			Set state = CreateObject("Scripting.Dictionary")
			InitData null
		Else
			LetState name, false
			LetState "lastError", ""
		End If
	End Function
	'-----
	Private Function GetData(name)
		If data.Exists(name) Then
			Set GetData = data.Item(name)
		Else
			Set GetData = Nothing
		End If
	End Function
	'-----
	Private Function SetData(name,value)
		If data.Exists(name) Then
			Set data.Item(name) = value
		Else
			data.Add name, value
		End If
	End Function
	'-----
	Private Sub InitData(name)
		If IsNull(name) Then
			Set data = CreateObject("Scripting.Dictionary")
		Else
			SetData name,CreateObject("Scripting.Dictionary")
		End If
	End Sub    
	'-----
	Private Sub AddItemsInData(name,value)
		GetData(name).Add value, value
	End Sub
	'-----
	Public Function Connect(ConfigFtpUpload)
		'�������� �� ���������� �� � ������� � ����������, ���� ��
		Disconnect        
		InitParamComponent ConfigFtpUpload
		If compFtp.Connect() = 1 Then
			LetState "connect", true
		End If
		Connect = GetState("connect")
	End Function
	'-----
	Public Sub Disconnect()
		If GetState("connect") Then
			compFtp.Disconnect()
			InitState null
		End If
	End Sub
	'-----
	Private Function LetParamCompFtp(ConfigFtpConnection,name)
		LetState "paramCompFtp" & name, GetParameterXML(ConfigFtpConnection,name)
		LetParamCompFtp = GetState("paramCompFtp" & name)
	End Function
	'-----
	Public Function GetParamFtp(name)
		GetParamCompFtp = GetState("paramCompFtp" & name)
	End Function
	'-----
	Private Sub InitParamComponent(ConfigFtpUpload)
		Dim ConfigFtpConnection
		
		'��� ����������� ������� � ������������ ����� ������ � ����������� ������, ������� �������� ��������� ������ ����������� ��� �� ������
		Set ConfigFtpConnection = ConfigXML.selectSingleNode("//*[@idName='" & ConfigFtpUpload.getAttribute("name") & "']")
		compFtp.Hostname = LetParamCompFtp(ConfigFtpConnection,"serverUri")
		compFtp.Port = LetParamCompFtp(ConfigFtpConnection,"serverPort")
		'������ ����� ���� �� �������� ����� � ������� �� ������, � ftp-������ ������� � ����, ������� �� ��������� ��������� �����,
		'�� ����� � ��������� � ������. �����
		compFtp.Passive = IIf(LetParamCompFtp(ConfigFtpConnection,"passiveMode") = "0",false,true)
		compFtp.Username = LetParamCompFtp(ConfigFtpConnection,"login")
		compFtp.Password = LetParamCompFtp(ConfigFtpConnection,"password")
	End Sub
	'-----
	Public Function ChangeRemoteDir(ConfigFtpUpload)    
		'�����/������������� ����������� ���������
		InitState("changeDir") 
		'�������� ������� ����������
		If compFtp.ChangeRemoteDir(GetParameterXML(ConfigFtpUpload,"folder")) = 1 Then
			'��������� ���������� �� ���
			If compFtp.GetCurrentRemoteDir() = GetParameterXML(ConfigFtpUpload,"folder") Then
				LetState "changeDir",true
			Else
				LetState "lastError",compFtp.LastErrorText
			End If
		End If
		ChangeRemoteDir = GetState("changeDir")
	End Function
	'-----
	Public Function ConnectAndChangeRemoteDir(ConfigFtpUpload)
		If Connect(ConfigFtpUpload) Then
			ChangeRemoteDir ConfigFtpUpload
		End If
		ConnectAndChangeRemoteDir = GetState("connect") And GetState("changeDir")
	End Function
	'-----    
	Public Function PutFile(localFileName,remoteFileName)    
		'�����/������������� ����������� ���������
		InitState("putFile") 
		'������������� ����
		If compFtp.PutFile(localFileName,remoteFileName) = 1 Then
			LetState "putFile",true
		Else
			LetState "lastError",compFtp.LastErrorText
		End If
		PutFile = GetState("putFile")
	End Function
	'-----    
	Public Function GetListFiles(pattern)
		Dim FilesListXMLStr, FilesListXML
		Dim NodeFile
		
		InitData("dirFileList")
		'�������� ������ ������ � XML �������
		FilesListXMLStr = compFtp.GetCurrentDirListing(pattern)
		'������ � ���������� ������ �����
		Set FilesListXML = CreateObject("MSXML2.DOMDocument.6.0")
		If FilesListXML.loadXML(FilesListXMLStr) Then
			LetState "dirListing", true
			For Each NodeFile In FilesListXML.documentElement.getElementsByTagName("file")
				AddItemsInData "dirFileList", GetParameterXML(NodeFile,"name")
			Next
		End If
		GetListFiles = GetData("dirFileList").Items
	End Function
	'-----
	Function DeleteFile(FileName)
		InitState("deleteFile") 
		'������������� ����
		If compFtp.DeleteRemoteFile(FileName) = 1 Then
			LetState "deleteFile",true
		Else
			LetState "lastError",compFtp.LastErrorText
		End If
		DeleteFile = GetState("deleteFile")        
	End Function
End Class
'------------------------------
Sub WriteToLog(Message,FolderLogFilePath)
	Const ForAppending = 8, TristateTrue = 0
	Dim tFilePath, tFile

	tFilePath = FolderLogFilePath & "\" & fso.GetBaseName(WScript.ScriptFullName) &".log"
	Set tFile = fso.OpenTextFile(tFilePath, ForAppending, true, TristateTrue)
	tFile.WriteLine("--" & Date & " " & FormatDateTime(Now, vbShortTime) & ": " & Trim(Message))
	tFile.Close
End Sub
'------------------------------
Sub SendError(Message)
	Dim FolderLogFilePath
	'��� ������� ����� �������� �� �����
	'MsgBox(Message)
	'� ������� ������ ����� � ���� [��� �������].log � ����� � ��������������� ������, ��� � ����� �� �������� ��� ��� ����������
	If IsEmpty(ConfigFilePath) Then
		FolderLogFilePath = CurrentScriptFolder.Path
	Else
		FolderLogFilePath = fso.GetParentFolderName(ConfigFilePath)
	End If
	WriteToLog Message, FolderLogFilePath
End Sub
'------------------------------
'����� ������ ��� �������
Sub SendDebug(Message)
    If Debug Then
        SendError "--- " & Message
    End If    
End Sub
'------------------------------
Sub Prepare
	'���������� �������� ������� (�����)
	Set fso = CreateObject("Scripting.FileSystemObject")
	'�������� �����, ��� ����� ���� �������
	Set CurrentScriptFolder = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))
	Set Shell = CreateObject("WScript.Shell")
End Sub
'------------------------------
Sub ParseXMLConfigFile(ConfigFilePath)
	Dim FileSchemas,xmlSchema
	'��������� ���������������� ���� �������� �� �������, �� ���� ��������� ��� �� ����� XML (���� ��� ����), ����� ����� �� ������ ������ � ����
	FileSchemas = CurrentScriptFolder.Path & "\config.xsd"
	If fso.FileExists(FileSchemas) Then 
		Set xmlSchema = CreateObject("MSXML2.XMLSchemaCache.6.0")
		xmlSchema.add "", FileSchemas
	End If

	'������ �� XML ����� �������� �� ���� �������������, �� ����������� �� �����
	Set ConfigXML = CreateObject("MSXML2.DOMDocument.6.0")
	If Not IsEmpty(xmlSchema) Then
		Set ConfigXML.schemas = xmlSchema
	End If
	ConfigXML.async="false"

	If Not ConfigXML.load(ConfigFilePath) Then
		SendError "!!!������ ��������� ����������������� ����� " & chr(34) & ConfigFilePath & chr(34) & vbCrLf & _
				  "Reason: " & ConfigXML.parseError.reason & vbCrLf & _
				  "Source: " & ConfigXML.parseError.srcText & vbCrLf & _
				  "Line: " & ConfigXML.parseError.Line & vbCrLf & _
				  "�������� ��������."
		Wscript.Quit
	End If
End Sub
'------------------------------
Sub LoadConfig
	'��������� ��������� ����� � ���������������� �����. �� ����� ���� ������� � ������ ��������� ��� ������� �������, ���� �� 
	'�� ��������� ���� ���� ������������ �� ����������� ������ config.xml � ���������� �������.
	If WScript.Arguments.Count > 0 Then
		ConfigFilePath = WScript.Arguments(0)
	Else
		ConfigFilePath = CurrentScriptFolder.Path & "\config.xml"
	End If

	If Not fso.FileExists(ConfigFilePath) Then
		SendError "!!!�� ������ ���������������� ���� " & chr(34) & ConfigFilePath & chr(34) & ". �������� ���� ��� ��������."
		WScript.Quit
	Else
		'����������� ���� � ����������������� ����� � ������ ���
		ConfigFilePath = fso.GetFile(ConfigFilePath).Path
		ParseXMLConfigFile(ConfigFilePath)
	End If
End Sub
'------------------------------
Sub CreateFolderRecursive(FullPath)
	Dim arr, dir, path
  
	arr = split(FullPath, "\")
	path = ""
	For Each dir In arr
		If Not path = "" Then
			path = path & "\"
		End If
		path = path & dir
		If Not fso.FolderExists(path) Then
			fso.CreateFolder(path)
		End If
	Next
End Sub
'------------------------------
Function GetTempUploadFolderPath
	'�����/�������� ��������� ���������� ��� �������� ��������
	Dim PathConfig

	If IsEmpty(TempUploadFolder) Then
		'�������� �� �������� ���� � ����, ���� ����� �������� ��������
		'� ������ ���������� ������� � ����� ������� ����� Temp, ������� ����� �������� ������
		'��������� ����� �� ����������� ������ � ��� (����� �� ������������ AD �������)
		PathConfig = GetParameterXML(GetNodeXMLRoot("parameters"),"tempUploadFolder")
		If PathConfig = "" Then
			PathConfig = CurrentScriptFolder.Path & "\Temp"
		End If

		'� �������� ���������/�������� ��������� ���������� ������������� ������
		On Error Resume Next
		CreateFolderRecursive PathConfig
		If Not Err.number = 0 Then
			SendError "!!!������ ��������� ��������� ����������. �������� ���� ��� ��������. " & Err.Description
			'����������� ������ - �������
			WScript.Quit
		End If
		Set TempUploadFolder = fso.GetFolder(PathConfig)
	End If
	Set GetTempUploadFolderPath = TempUploadFolder
End Function
'------------------------------
Function Is64BitSystem
	Is64BitSystem = (right(Shell.environment("system").item("processor_architecture"), 2) = "64")
End Function
'------------------------------
Function IIf(expr, truepart, falsepart)
	If expr Then
		IIf = truepart
	Else
		IIf = falsepart
	End If
End Function
'------------------------------
Function GetNodeXMLRoot(NodeName)
	'��� �������� ��������� �������� ����� ����. �����
	Set GetNodeXMLRoot = GetNodeXML(ConfigXML.documentElement,NodeName)
End Function
'------------------------------
Function GetParameterXML(DOMElement,Parameter)
	Dim NodeXML
	Set NodeXML = GetNodeXML(DOMElement,Parameter)
	If NodeXML Is Nothing Then
		GetParameterXML = ""
	Else
		If NodeXML.childNodes.length = 0 Then
			GetParameterXML = ""
		Else
			GetParameterXML = NodeXML.childNodes(0).nodeValue
		End If
	End If
End Function
'------------------------------
Function GetNodeXML(DOMElement,NodeName)
	'��������� ������� ����������� ���� � ���������� ��������
	Dim DOMList
	Set DOMList = DOMElement.getElementsByTagName(NodeName)
	If Not DOMList.length = 0 Then
		Set GetNodeXML = DOMList.item(0)
	Else
		Set GetNodeXML = Nothing
	End If
End Function
'------------------------------
Function FindCore1C
	Dim CLSID_1C8Pointer
	'����� �� �������� � � ��� �������� ����������� ���� 1�
	'Core1CPath = "C:\Program Files\1cv82\common\1cestart.exe"
	'�� ��������� �� ������ ����� �� ����, ������ �������������, 1� ����� ���� ����������� � ������ �������, ��� ������ ������� �64, �� ���� � ������� ������ �� �����������
	'!!!�� ���������� �� �64 �������� � �64 �������� ���������� �� ������� ���������� ������������� ����� ������ ����
    '� ������� ���������� ������������ ������� �������� "version", ������� ��������� ������������ ������ 1� (����� ���� 82 ��� 83)
	On Error Resume Next
	CLSID_1C8Pointer = Shell.RegRead("HKEY_CLASSES_ROOT\V" & GetParameterXML(GetNodeXMLRoot("parameters"),"version") & ".Application\CLSID\")
	Core1CPath = Shell.RegRead("HKEY_CLASSES_ROOT\" & IIf(Is64BitSystem,"Wow6432Node\","") & "CLSID\" & CLSID_1C8Pointer & "\LocalServer32\")
	If Not Err.number = 0 Then
		SendError "!!!������ ������ � ������� ���� � ������������ ����� 1�. " & Err.Description
		SendError "!!!�� ������� ��� ����������� ���������/���������� ��� ������� ��������. ���������� � �������� ���� ��� ��������."
		WScript.Quit
	End If
End Function
'------------------------------
Sub ConnectComponent(Var,NameComponent,ComponentDLLPath)
	Dim Component, resultRegister

	On Error Resume Next
	Set Component = CreateObject(NameComponent)
	If Not Err.number = 0 Then
		'�� ��������������� ���������
		If Err.number = 429 Then
			SendError "!������ " & Err.number & ". ���������� " & chr(34) & NameComponent & chr(34) & " �� ���������������� � �������."
		ElseIf Err.number = -2147024770 Then
			SendError "!������ " & Err.number & ". �������/��������� ���� dll ����� ������������������ ���������� " & chr(34) & NameComponent & chr(34)
		Else
			SendError "!������ " & Err.number & ". �� ������ ���������� ���������� " & chr(34) & NameComponent & chr(34) & ". " & Err.Description
		End If
		Err.Clear
		'������� �������� ���������� ��� ������� dll
		If fso.FileExists(ComponentDLLPath) Then
			'�� ������ ������ ������� ������������� ���������� - ����� ��� ������ �������� (������� dll, � ������ � �������� ��������, ��� ������ �����������)
			Set resultRegister = Shell.Exec("regsvr32.exe /s /u " & chr(34) & ComponentDLLPath & chr(34))
			If Not Err.number = 0 Then
				SendError "!������ ��� �������� ���������� " & chr(34) & NameComponent & chr(34) &" �� ����� " & chr(34) & ComponentDLLPath & chr(34) & ". " & Err.Description
				Err.Clear
			Else
				If Not IsEmpty(resultRegister) Then
					'���� ���� ���������� ��������
					Do While resultRegister.Status = 0
						WScript.Sleep 100
					Loop
				End If
			End If
			'������������ � ������� dll-�� � ����� ������
			Set resultRegister = Shell.Exec("regsvr32.exe /s " & chr(34) & ComponentDLLPath & chr(34))
			If Not Err.number = 0 Then
				SendError "!������ ��� ����������� ���������� " & chr(34) & NameComponent & chr(34) & " �� ����� " & chr(34) & ComponentDLLPath & chr(34) & ". " & Err.Description
			Else 
				If Not IsEmpty(resultRegister) Then
					'���� ���� ���������������� ����������
					Do While resultRegister.Status = 0
						WScript.Sleep 100
					Loop
				End If
				SendError "���������� � ����� ������ ���������� " & chr(34) & NameComponent & chr(34) & " �� ����� " & chr(34) & ComponentDLLPath & chr(34) & ". " & Err.Description
				'������� ����� �������� ���������. ���� ������, �� � ��� ��� �������� � ������ ������
				Set Component = CreateObject(NameComponent)
				If Not Err.number = 0 Then
					SendError "!������ ����������� ���������� " & chr(34) & NameComponent & chr(34) & " ����� ���������. " & Err.Description
					SendError "!��������� ������������ ������� �������. C�. ���� readme.txt"
				Else
					Set Var = Component
				End If
			End If
		Else
			SendError "!�� ������ ���� " & chr(34) & ComponentDLLPath & chr(34) & " ��� ����������� ���������� " & chr(34) & NameComponent & chr(34)
		End If
	Else
		Set Var = Component
	End If
End Sub
'------------------------------
Function GetPresentActivityClientServerBases()
	Dim BaseConfig
	GetPresentActivityClientServerBases = False
	For Each BaseConfig in GetNodeXMLRoot("bases").childNodes
		If Not BaseConfig.getAttribute("activity") = "0" And Not GetNodeXML(BaseConfig,"clientServer") Is Nothing Then
			GetPresentActivityClientServerBases = True
		End If
	Next
End Function
'------------------------------
Function GetPresentActivityFtpUploadBases()
	Dim BaseConfig, FtpUploadNode
	GetPresentActivityFtpUploadBases = False
	For Each BaseConfig in GetNodeXMLRoot("bases").childNodes
		If Not BaseConfig.getAttribute("activity") = "0" Then
			Set FtpUploadNode = GetNodeXML(BaseConfig,"ftpUpload")
			If Not FtpUploadNode Is Nothing Then
				If Not FtpUploadNode.getAttribute("activity") = "0" Then
					GetPresentActivityFtpUploadBases = True
				End If
			End If
		End If
	Next
End Function
'------------------------------
Function ConnectComConnector1C8
	'COM-��������� ����� � ����� � ����������� ������ 1�
	'���������� �����, ����� ���� �������� �� �������� ������-��������� ����
	If GetPresentActivityClientServerBases() Then
		ConnectComponent ComConnector8, "V" & GetParameterXML(GetNodeXMLRoot("parameters"),"version") & ".COMConnector", fso.GetParentFolderName(Core1CPath) & "\comcntr.dll"
		If IsEmpty(ComConnector8) Then
			SendError "!!�������� ��� � ������� ���������� ����� ���������, �.�. �� ��������� COM-��������� 1� " & GetParameterXML(GetNodeXMLRoot("parameters"),"version") & "."
		End If
	End If
End Function
'------------------------------
Function ConnectFtpComponent
	Dim ComponentDLLPath, Component
	'� �������� ���������� ��� ������ � ftp ���������� ChilkatFtp. ��� freeware, ����� ������������� � ������� � ���������� ��� ������� ��������.
	'� ����� ������������� http://www.chilkatsoft.com/downloads.asp ��� ��� ������, ������� dll � help-���� � ��������� ����� � ����� ftp �����������
	'�� ����� ���������� � ������ ����� (� ��������� �����������) - ���� � dll ��������� � ������-����� � ����� root/parameters/ftpComponentPathDLL
	'P.S. ���������� � ������� ftp.exe �� ��������� ������ ������������� � �������� ������ � ���
	'����������� ���������� ������������ ������ ���� ���� ��������� ���-����������� � ���������������� �����, � ����� ���� � ��������� ����������� �������� �� ���
	If Not (GetNodeXMLRoot("ftpConnections") Is Nothing) And GetPresentActivityFtpUploadBases() Then
		ComponentDLLPath = GetParameterXML(GetNodeXMLRoot("parameters"),"ftpComponentPathDLL")
		ComponentDLLPath = IIf(ComponentDLLPath = "",CurrentScriptFolder.Path & "\ftp\ChilkatFTP.dll",ComponentDLLPath)
		ConnectComponent Component, "ChilkatFTP.ChilkatFTP", ComponentDLLPath
		If IsEmpty(Component) Then
			SendError "!!�������� ������ �� ftp-������ �� ����� ��������. �������� ��� ����� ������������ �� ��������� ����������."
		Else
			Set Ftp = New FtpClass
			Set Ftp.Component = Component
		End If
	End If
End Function
'------------------------------
'��������� ������ � ������������� ������, ���������� ��� ������
Function AddInFixedArray(VarArray,VarData)
    Dim CurrentSizeVarArray
    
    CurrentSizeVarArray = UBound(VarArray,1)
	'���� ������� ������� �� ������ - ����� ����������� ������
    If Not IsEmpty(VarArray(0)) Then
		'����������� ������ ������������� ������� �� 1
        CurrentSizeVarArray = CurrentSizeVarArray + 1
		ReDim Preserve VarArray(CurrentSizeVarArray)
	End If
    Set VarArray(CurrentSizeVarArray) = VarData
End Function
'------------------------------
Function FindBaseOnClientServer(dicUploadData)
	Dim Cluster
	'������� �������
    Dim WorkingProcess
	'������ ���� ������� ���������
    ReDim ConnectWorkingProcesses(0)
    '������� ������� �������
    Dim ConnectWorkingProcess
    '����, ������� ���� � ���������
	Dim TempBase

	'������������� � ������ �������� 1�
	On Error Resume Next
	
	If Err.number = 0 Then
		On Error Goto 0
		
        Set Cluster = dicUploadData.Item("COMAgent1C").GetClusters()(0)
        '������� ��� �������������� �������� �� ������� � ������� - ���������� ������ ������
		'����� ����� �������� ��������� �������������� �������� � ��������������� ���� � ��� ���������
		dicUploadData.Item("COMAgent1C").Authenticate Cluster, "", ""
		
		For Each WorkingProcess in dicUploadData.Item("COMAgent1C").GetWorkingProcesses(Cluster)
            '������� - �� ������ ������� �� 1�: ������� ������� � ���������� � ������� ���������
            '������ ���������� � ������� ��������� ����� ������() ��� ������ � ���
			Set ConnectWorkingProcess = ComConnector8.ConnectWorkingProcess("tcp://"+dicUploadData.Item("BaseConnection").Item("serverName")+":"+CStr(WorkingProcess.MainPort))		
            '������ ConnectWorkingProcess ����� ��� ����� ��� ����������� �������������
			'����� �����������. ����� ����� ������ ������ �������������
            ConnectWorkingProcess.AddAuthentication dicUploadData.Item("BaseConnection").Item("login"), dicUploadData.Item("BaseConnection").Item("password")
	    
            AddInFixedArray ConnectWorkingProcesses, ConnectWorkingProcess
		Next
        dicUploadData.Add "ConnectWorkingProcessesClientServer", ConnectWorkingProcesses

        '���������� ���� � ����� �� ������� ���������, � ���� ������ ���
        For Each TempBase in ConnectWorkingProcesses(0).GetInfoBases()
			If TempBase.name = dicUploadData.Item("BaseConnection").Item("baseName") Then
				'����� ������������� ����, ��������� � ����� ������
				dicUploadData.Add "BaseClientServer", TempBase
				Exit For
			End If
		Next
	Else
		SendError "!!������ ����������� � ������� ���������� ��� ���� " & FormatBaseNameLog(dicUploadData) & ". �������� ��������� ���� ��������. �������� ������: " & Err.Description
		Err.Clear
	End If
	FindBaseOnClientServer = dicUploadData.Exists("BaseClientServer")
End Function
'------------------------------
Function DisconnectUsers(dicUploadData)
	Dim CurrentConnections, CurrentConnection
    Dim ConnectWorkingProcess

	DisconnectUsers = True
	'������� ��� �������� ����������, ���� ��� �� � ����������� ������ (� ����������� �������� ������������ � ���-����������)...
	On Error Resume Next
	
    For Each ConnectWorkingProcess in dicUploadData.Item("ConnectWorkingProcessesClientServer")
        CurrentConnections = ConnectWorkingProcess.GetInfoBaseConnections(dicUploadData.Item("BaseClientServer"))
	    If Not Err.number = 0 Then
		    SendError "!!������ ��������� ������� ���������� ���� " & FormatBaseNameLog(dicUploadData) & ". ��������� ��������� ��������������. " & Err.Description
		    Err.Clear
		    DisconnectUsers = False
	    Else
		    For Each CurrentConnection in CurrentConnections
			    '��������� ������ ������� �������. ����������������, ������� � ������� Com-���������� �� �������
			    If CurrentConnection.AppID = "1CV8" Then
				    If CurrentConnection.IBConnMode = 0 Then
					    ConnectWorkingProcess.Disconnect(CurrentConnection)
				    Else
					    '���� ���������� ����������� � ����������� ������ - �������� ����������, ���������� ��������
                        SendError "!������� ����������� ����������, �������� ���������"
					    DisconnectUsers = False
				    End If
			    ElseIf CurrentConnection.AppID = "Designer" Then
				    '������ ������������ - ���������� �������� (���� ����� � ������� ���! ��� ����� ��������� ��� ���������� ����� �� ����������)
                    SendError "!������ ������������, �������� ���������"
				    DisconnectUsers = False
			    End If
		    Next
	    End If
    Next
	dicUploadData.Add "DisconnectUsers", DisconnectUsers
End Function
'------------------------------
Function BlockBase(dicUploadData,mode)
    Dim Base, ConnectWorkingProcess
    
    Set Base = dicUploadData.Item("BaseClientServer")
    Base.ScheduledJobsDenied = mode
    Base.SessionsDenied = mode

    Set ConnectWorkingProcess = dicUploadData.Item("ConnectWorkingProcessesClientServer")(0)
    ConnectWorkingProcess.UpdateInfoBase(Base)
End Function
'------------------------------
Function ReadTextFileUTF8(FilePath)
    Dim objStream, strData

    Set objStream = CreateObject("ADODB.Stream")

    objStream.CharSet = "utf-8"
    objStream.Open
    objStream.LoadFromFile(FilePath)

    ReadTextFileUTF8 = objStream.ReadText()

    objStream.Close
    Set objStream = Nothing
End Function
'------------------------------
Function GetFolderUploadPath(dicUploadData)
	On Error Resume Next
	'���� ������� ����� �������� � ���������� ���� - �� ��������� �� �������/������ � ��������� � ���, ����� �� ��������� �����
	GetFolderUploadPath = GetParameterXML(dicUploadData.Item("BaseConfig"),"baseUploadFolder")
	IF Not GetFolderUploadPath = "" Then
		CreateFolderRecursive GetFolderUploadPath
		If Not Err.number = 0 Then
			'���� �������� ��������� � �������� � ����� �������� - ����� ������ � ���������� �������� �� ��������� �����
			SendError "!������ �������/�������� ����� " & chr(34) & GetFolderUploadPath & chr(34) & " ��� ���� " & FormatBaseNameLog(dicUploadData) & _
					  ". ���������� �������� �� ��������� �����. �������� ������: " & Err.Description
			Err.Clear
			GetFolderUploadPath = GetTempUploadFolderPath
		End If
	Else
		GetFolderUploadPath = GetTempUploadFolderPath
	End If
End Function
'------------------------------
Function GetFileUploadPath(dicUploadData)
	'����� � ��������� ���������, ����� ����� ���� �������������� ������ ����� � ����������������� ��� ��������
	Dim fileNamePrefixUpload
	
	If dicUploadData.Item("BaseConnection").Item("type") = "clientServer" Then
		fileNamePrefixUpload = IIf(GetParameterXML(dicUploadData.Item("BaseConfig"),"fileNamePrefixUpload") = "",dicUploadData.Item("BaseConnection").Item("baseName"),GetParameterXML(dicUploadData.Item("BaseConfig"),"fileNamePrefixUpload"))
		dicUploadData.Add "fileNamePrefixUpload", fileNamePrefixUpload
	End If
	GetFileUploadPath = fileNamePrefixUpload & "_" & Year(Date) & "_" & Right("0" & Month(Date), 2) & "_" & Right("0" & Day(Date), 2) & ".dt"
End Function
'------------------------------
Function FormatBaseNameLog(dicUploadData)
	If dicUploadData.Item("BaseConnection").Item("type") = "clientServer" Then
		FormatBaseNameLog = chr(34) & dicUploadData.Item("BaseConnection").Item("serverName") & "\" & dicUploadData.Item("BaseConnection").Item("baseName") & chr(34)
	End If
End Function
'------------------------------
Function FtpUpload(dicUploadData)
	'���� �������� ��� ��� ����������� ������ ��� ����������� � ����������
	FtpUpload = false
	If Ftp.Connect(dicUploadData.Item("ConfigFtpUpload")) Then
		If Ftp.ChangeRemoteDir(dicUploadData.Item("ConfigFtpUpload")) Then
			If Ftp.PutFile(dicUploadData.Item("FileUploadPath"),fso.GetFileName(dicUploadData.Item("FileUploadPath"))) Then
				SendError "����������� �������� " & chr(34) & dicUploadData.Item("FileUploadPath") & chr(34) & " �� ftp-������ � ������. ������ " & chr(34) & dicUploadData.Item("ConfigFtpUpload").getAttribute("name") & chr(34) & "."
				FtpUpload = true
			Else
				SendError "!!�� ������ �������� ���� �������� �� ftp-������ " & chr(34) & Ftp.GetParamFtp("serverUri") & chr(34) & ". ���������� �������� "& chr(34) & _
						  dicUploadData.Item("FileUploadPath") & chr(34) & " �� ftp-������ ��������. " & vbCrLf & Ftp.GetState("lastError")
			End If
		Else
			SendError "!!�� ������ �������� ������� ���������� �� ftp-������� " & chr(34) & Ftp.GetParamFtp("serverUri") & chr(34) & ". ���������� �������� " & chr(34) & _
				  dicUploadData.Item("FileUploadPath") & chr(34) & " �� ftp-������ ��������. " & vbCrLf & Ftp.GetState("lastError")
		End If
	Else
		SendError "!!������ ����������� � ftp-������� � ������. ������ " & chr(34) & dicUploadData.Item("ConfigFtpUpload").getAttribute("name") & chr(34) & ". ���������� �������� " & chr(34) & _
				  dicUploadData.Item("FileUploadPath") & chr(34) & " �� ftp-������ ��������. " & vbCrLf & Ftp.GetState("lastError")
	End If
	Ftp.Disconnect()
End Function
'------------------------------
Function UploadBaseFtp(dicUploadData)
	UploadBaseFtp = False
	If Not IsEmpty(Ftp) Then
		UploadBaseFtp = FtpUpload(dicUploadData)
		If UploadBaseFtp Then
			'����� �� ������� ���� �� ��������� ���������� ����� ����������. �� ��������� �������.
			If Not dicUploadData.Item("ConfigFtpUpload").getAttribute("deleteAfterUpload") = "0" Then
				fso.DeleteFile(dicUploadData.Item("FileUploadPath"))
			End If
		End If
	Else
		SendError "!!Ftp ���������� �� ����������. ���������� �������� " & chr(34) & dicUploadData.Item("FileUploadPath") & chr(34) & " �� ftp-������ ��������."
	End If
End Function
'------------------------------
Sub TransferToFtp(dicUploadData)
	Dim ConfigFtpUpload
	
	Set ConfigFtpUpload = GetNodeXML(dicUploadData.Item("BaseConfig"),"ftpUpload")
	If Not (ConfigFtpUpload Is Nothing) And dicUploadData.Item("ResultUploadBaseDT") Then
		If Not ConfigFtpUpload.getAttribute("activity") = "0" Then
			dicUploadData.Add "ConfigFtpUpload", ConfigFtpUpload
			dicUploadData.Item("ResultUploadBaseFtp") = UploadBaseFtp(dicUploadData)
		End If
	End If
End Sub
'------------------------------
Sub UploadBaseClientServer2(dicUploadData)
	Dim CommandStringUpload, resultUpload
    Dim FileDumpResult, DumpResult, FileOutResult

	'����� �������� ���������� ��� ��������� ������ ��������
    '��� �������� �������� ������� ���� ������ ���������� �� �������� (��� ������� � ������ ���������� � ������ ������)
    '� ����� ���� ����� ��� �������� ������������ �������� (��� ����� 0,1, 101)
    dicUploadData.Add "FolderUploadPath", GetFolderUploadPath(dicUploadData)
	dicUploadData.Add "FileUploadPath", dicUploadData.Item("FolderUploadPath") & "\" & GetFileUploadPath(dicUploadData)
    dicUploadData.Add "OutResultPath", dicUploadData.Item("FolderUploadPath") & "\out.log"
    dicUploadData.Add "DumpResultPath", dicUploadData.Item("FolderUploadPath") & "\dump.log"

	CommandStringUpload = chr(34) & Core1CPath & chr(34) & " config /S" & dicUploadData.Item("BaseConnection").Item("serverName") & "\" & dicUploadData.Item("BaseConnection").Item("baseName") & _
                          " /n" & chr(34) & dicUploadData.Item("BaseConnection").Item("login") & chr(34) & " /p" & chr(34) & dicUploadData.Item("BaseConnection").Item("password") & chr(34) & _
                          " /DumpIB" & dicUploadData.Item("FileUploadPath") & " /Out" & dicUploadData.Item("OutResultPath") & " /DumpResult" & dicUploadData.Item("DumpResultPath") & _
                          " /UC" & chr(34) & dicUploadData.Item("BaseClientServer").PermissionCode & chr(34)
                          
	On Error Resume Next

	Set resultUpload = Shell.Exec(CommandStringUpload)
	If Err.number = 0 Then
		Do While resultUpload.Status = 0
			WScript.Sleep 200
		Loop
		On Error Goto 0
		'����������� ������� ��������, �������� ����� ������ � �����
        DumpResult = ReadTextFileUTF8(dicUploadData.Item("DumpResultPath"))
        If DumpResult = "0" Then
            dicUploadData.Item("ResultUploadBaseDT") = True
		    SendError "��������� ���� " & FormatBaseNameLog(dicUploadData) & " � " & chr(34) & dicUploadData.Item("FileUploadPath") & chr(34)
        Else
           Set FileOutResult = fso.OpenTextFile(dicUploadData.Item("OutResultPath"), 1, False) 
           SendError Trim("!!������ ��� �������� ���� " & FormatBaseNameLog(dicUploadData) & ": " & FileOutResult.ReadAll)
           FileOutResult.Close
        End If
                
        WScript.Sleep 1000
        '�������� ����� �����
        fso.DeleteFile dicUploadData.Item("DumpResultPath")
        fso.DeleteFile dicUploadData.Item("OutResultPath")
	End If	
End Sub
'------------------------------
Sub UploadBaseClientServer(dicUploadData)
	ConnectAgent1C dicUploadData
    
    If FindBaseOnClientServer(dicUploadData) Then
		'������ ���������� �� ����
        BlockBase dicUploadData, true
        '�������� ������������� � ����
		If DisconnectUsers(dicUploadData) Then
			'���� ���� ������� - ���������� ��������. ���������� ���� ��� �������� �� ������, �.�. ������ �� �������� ��������, ���� ����������� ����� � ��� �������� ���� �����������.
			UploadBaseClientServer2 dicUploadData
			If Not dicUploadData.Item("ResultUploadBaseDT") Then
				SendError "!!�� ������ ��������� ���� " & FormatBaseNameLog(dicUploadData) & " ����� ������ ��������� ������ � �����������."
			End If
		Else
			SendError "!!�� ������ ��������� ���� ������������� ��� ���� " & FormatBaseNameLog(dicUploadData) & ". �������� ��������� ���� ��������."
		End If
        '������� ���������� � ����
        BlockBase dicUploadData, false
	Else
		SendError "!!�� ������� �� ������� ���� " & FormatBaseNameLog(dicUploadData) & ". �������� ��������� ���� ��������."
	End If
    CloseConnectTo1C dicUploadData
End Sub
'------------------------------
Sub ConnectAgent1C(dicUploadData)
    Dim COMAgent1C

    SendDebug "start connect COM agent 1C"
    Set COMAgent1C = ComConnector8.ConnectAgent("tcp://"+dicUploadData.Item("BaseConnection").Item("serverName"))
    dicUploadData.Add "COMAgent1C", COMAgent1C
    SendDebug "end connect COM agent 1C"
End Sub
'------------------------------
'��������� ��� �������� ����������� � 1� - ����� � ���������� � �������� ����������
Sub CloseConnectTo1C(dicUploadData)
    Dim ConnectWorkingProcess

    SendDebug "start disconnect base 1C"
    Set dicUploadData.Item("BaseClientServer") = nothing
    SendDebug "end disconnect base 1C"
    SendDebug "start disconnect connection working processes"
    For Each ConnectWorkingProcess in dicUploadData.Item("ConnectWorkingProcessesClientServer")
        Set ConnectWorkingProcess = nothing
    Next
    SendDebug "end disconnect connection working processes"
    SendDebug "start disconnect COM agent 1C"
    Set dicUploadData.Item("COMAgent1C") = nothing
    SendDebug "end disconnect COM agent 1C"
    
    WScript.Sleep 1000
End Sub
'------------------------------
Sub RestructuringFilesAdd(dicUploadData,File)
	Dim ArrDate
	Dim lenPrefixName, tempStrucYear
	
	If Not dicUploadData.Exists("RestructuringFilesTree") Then
		dicUploadData.Add "RestructuringFilesTree", CreateObject("Scripting.Dictionary")
	End If
	If Not dicUploadData.Exists("RestructuringFiles") Then
		dicUploadData.Add "RestructuringFiles", CreateObject("Scripting.Dictionary")
	End If
	'��������� �������, ����� �� �������� ����� �����. ���������� ����� ����� ���� �����.
	lenPrefixName = Len(dicUploadData.Item("fileNamePrefixUpload"))
	If Left(fso.GetFileName(File),lenPrefixName) = dicUploadData.Item("fileNamePrefixUpload") Then
		'�������� ������� � ������ ���� �����
		ArrDate = split(Mid(fso.GetBaseName(File),lenPrefixName+2),"_")
		If UBound(ArrDate) = 2 Then
			dicUploadData.Item("RestructuringFiles").Add fso.GetFileName(File), File
			If Not dicUploadData.Item("RestructuringFilesTree").Exists(ArrDate(0)) Then
				dicUploadData.Item("RestructuringFilesTree").Add ArrDate(0), CreateObject("Scripting.Dictionary")
			End If
			Set tempStrucYear = dicUploadData.Item("RestructuringFilesTree").Item(ArrDate(0))
			If Not tempStrucYear.Exists(ArrDate(1)) Then
				tempStrucYear.Add ArrDate(1), CreateObject("Scripting.Dictionary")
			End If
			tempStrucYear.Item(ArrDate(1)).Add ArrDate(2), fso.GetFileName(File)
		End If
	End If
End Sub
'------------------------------
Function RestructuringFilesMonthCheck(tempYear,tempMonth,depthInMonth)
	'������� ����� ������� ����� � ������� ������������ ������
	RestructuringFilesMonthCheck = (depthInMonth = 0) OR (DateDiff("m","01/" & tempMonth & "/" & tempYear,date) < depthInMonth)
End Function
'------------------------------
Sub RestructuringFilesSelect(dicUploadData)
	Dim PatternElement, rangeInDay
	Dim ArrYears, tempYear, ArrMonth, tempMonth, ArrDay, tempDay, tempDate, lastDate, i, FileTempPath

	'�������� �����, ������� ���������
	dicUploadData.Add "RestructuringFilesLeave", CreateObject("Scripting.Dictionary")
	For Each PatternElement In dicUploadData.Item("restructuringPattern").childNodes
		ArrYears = dicUploadData.Item("RestructuringFilesTree").Keys
		For Each tempYear In ArrYears
			ArrMonth = dicUploadData.Item("RestructuringFilesTree").Item(tempYear).Keys
			For Each tempMonth In ArrMonth
				'���������, �������� �� ��� ������� ����� �� �������
				If RestructuringFilesMonthCheck(tempYear,tempMonth,CInt(GetParameterXML(PatternElement,"depthInMonth"))) Then
					rangeInDay = CInt(GetParameterXML(PatternElement,"rangeInDay"))
					'��� ������ ������ ��������� � �������� �������, � ����� ������
					lastDate = DateSerial(tempYear,CInt(tempMonth)+1,0)
					lastDate = DateAdd("d",rangeInDay,IIF(lastDate < date,lastDate,date))
					ArrDay = dicUploadData.Item("RestructuringFilesTree").Item(tempYear).Item(tempMonth).Keys
					For i = 0 To UBound(ArrDay)
						tempDay = ArrDay(UBound(ArrDay)-i)
						tempDate = DateSerial(tempYear,tempMonth,tempDay)
						If i=0 OR rangeInDay = 0 Or DateDiff("d",tempDate,lastDate) > rangeInDay-1 Then
							lastDate = tempDate
							FileTempPath = dicUploadData.Item("RestructuringFilesTree").Item(tempYear).Item(tempMonth).Item(tempDay)
							If Not dicUploadData.Item("RestructuringFilesLeave").Exists(FileTempPath) Then
								dicUploadData.Item("RestructuringFilesLeave").Add FileTempPath, tempDate
							End If
						End If
					Next                    
				End If
			Next
		Next
	Next
	'�������� �����, ������� �������
	dicUploadData.Add "RestructuringFilesDelete", CreateObject("Scripting.Dictionary")
	For Each FileTempPath In dicUploadData.Item("RestructuringFiles").Keys
		If Not dicUploadData.Item("RestructuringFilesLeave").Exists(FileTempPath) Then
			dicUploadData.Item("RestructuringFilesDelete").Add FileTempPath, dicUploadData.Item("RestructuringFiles").Item(FileTempPath)
		End If
	Next
End Sub
'------------------------------
Sub RestructuringFtpDT(dicUploadData)
	Dim CurrentFile

	If Ftp.ConnectAndChangeRemoteDir(dicUploadData.Item("ConfigFtpUpload")) Then
		'�������� ��� ����� � ����� ��������, � ������� � ���������-������ � ��������� �� �����/�������
		For Each CurrentFile In Ftp.GetListFiles("*.*") 
			RestructuringFilesAdd dicUploadData, CurrentFile
		Next
		'��������� ������ ������, ������� �������
		RestructuringFilesSelect dicUploadData
		'������� �����
		For Each CurrentFile In dicUploadData.Item("RestructuringFilesDelete").Items
			Ftp.DeleteFile CurrentFile
		Next
	End If
End Sub
'------------------------------
Sub RestructuringLocalDT(dicUploadData)
	Dim CurrentFile
	'�������� ��� ����� � ��������� ����� ��������, � ������� � ���������-������ � ��������� �� �����/�������
	For Each CurrentFile In fso.GetFolder(dicUploadData.Item("FolderUploadPath")).Files 
		RestructuringFilesAdd dicUploadData, CurrentFile
	Next
	'��������� ������ ������, ������� �������
	RestructuringFilesSelect dicUploadData
	'������� �����
	For Each CurrentFile In dicUploadData.Item("RestructuringFilesDelete").Items
		CurrentFile.Delete()
	Next
End Sub
'------------------------------
Sub RestructuringBase(dicUploadData)
	Dim RestructuringLink

    SendDebug "start restructuring base"
	Set RestructuringLink = GetNodeXML(dicUploadData.Item("BaseConfig"),"restructuring")
	If Not RestructuringLink Is Nothing Then
		'�������� ������� ������������������ ������
		dicUploadData.Add "restructuringPattern", ConfigXML.selectSingleNode("//*[@idName='" & RestructuringLink.getAttribute("patternName") & "']") 
		'�������� ������ ������, ������ �� �� ��������� ������ ���, ������� ��������, � ������������ �������
		If dicUploadData.Item("ResultUploadBaseDT") Then		
			If dicUploadData.Item("ResultUploadBaseFtp") = True Then
				'������ ����� �� ���-�������
				RestructuringFtpDT dicUploadData
				SendError "������� ���������������� ����� " & chr(34) & GetParameterXML(dicUploadData.Item("ConfigFtpUpload"),"folder") & chr(34) & _
						  " ftp-������� � ������. ������ " & chr(34) & dicUploadData.Item("ConfigFtpUpload").getAttribute("name") & chr(34)
			Else
				'��������� �����
				RestructuringLocalDT dicUploadData
				SendError "������� ���������������� ��������� ����� " & chr(34) & dicUploadData.Item("FolderUploadPath") & chr(34)
			End If
		End If
	End If
    SendDebug "end restructuring base"
End Sub
'------------------------------
Sub UploadBase(dicUploadData)
	Dim dicBaseConnection
	Dim BaseConnection, BaseConnectionType
	
    SendDebug "start upload base"
	dicUploadData.Item("ResultUploadBaseDT") = False
	Set BaseConnection = GetNodeXML(dicUploadData.Item("BaseConfig"),"connection")
	'������� ��� ��������� ����������� � ���� � �������������� ����
	Set dicBaseConnection = CreateObject("Scripting.Dictionary")
	dicUploadData.Add "BaseConnection", dicBaseConnection
	dicBaseConnection.Add "login", GetParameterXML(BaseConnection,"login")
	dicBaseConnection.Add "password", GetParameterXML(BaseConnection,"password")
	'��������� ��� ����������� ���� 1� � ���������� ��������
	Set BaseConnectionType = GetNodeXML(BaseConnection,"type").childNodes(0)
	dicBaseConnection.Add "type", BaseConnectionType.nodeName
	If dicBaseConnection.Item("type") = "clientServer" Then
		dicBaseConnection.Add "serverName", GetParameterXML(BaseConnectionType,"serverName")
		dicBaseConnection.Add "baseName", GetParameterXML(BaseConnectionType,"baseName")
		UploadBaseClientServer dicUploadData
	ElseIf dicBaseConnection.Item("type") = "file" Then
		dicBaseConnection.Add "baseLocation", GetParameterXML(BaseConnectionType,"baseLocation")
	End If
    SendDebug "end upload base"
End Sub
'------------------------------
Sub UploadBases
	Dim dicUploadData
	Dim BaseConfig
    Dim KeysDict, i

	'���������� ����� �������
    Debug = (GetParameterXML(GetNodeXMLRoot("parameters"),"debug") = "1")
    '���������� ����� � ������ �� ����������������� ����� � �� ������� �� ���������
	For Each BaseConfig in GetNodeXMLRoot("bases").childNodes
		'������� ���������� ��� �����/�������� ������ ������� ���� ����� �����������
		SendDebug "start create dictionary"
        Set dicUploadData = CreateObject("Scripting.Dictionary")
		SendDebug "end create dictionary"
        dicUploadData.Add "BaseConfig", BaseConfig
        
		'���� ����� ������� �� ����� �������� ��������� <activity> = 0, ��� ������ �������� � ���������� �������� ������������
		SendDebug "start check activity base"
        If Not BaseConfig.getAttribute("activity") = "0" Then
			SendDebug "end check activity base"
            '���������� �������� ����
			UploadBase dicUploadData            
			'������������� �� ���-������ ��� �������������
			TransferToFtp dicUploadData
			'�������� ��������� ����� � ����������
			RestructuringBase dicUploadData
		End If
        
        SendDebug "start clear dictionary"
        Set dicUploadData = nothing
		SendDebug "end clear dictionary"
	Next
End Sub
'------------------------------
Sub Run
    Prepare	'���������� ����� ����������
    LoadConfig '�������� ����������������� �����
    FindCore1C '����� ���� 1�
    ConnectComConnector1C8 '����������� ���������� COM-���������� � 1� 8
    ConnectFtpComponent '����������� ���������� ��� ������ � ftp (��� �������������)
    UploadBases

    SendError "--------------------------------------"
End Sub
'------------------------------
Run