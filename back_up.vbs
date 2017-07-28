Option Explicit

Dim fso	'объект для работы с файловой системой - Scripting.FileSystemObject
DIm Shell 'объект WScript.Shell
Dim ConfigFilePath 'путь к конфигурационному файлу
Dim ConfigXML 'загруженный и проверенный (при наличии схемы) конф. файл - MSXML2.DOMDocument
Dim CurrentScriptFolder 'директория, где расположен скрипт
Dim Debug 'режим записи в лог детальных сообщений
Dim TempUploadFolder 'временная директория для помещения в нее выгрузок
Dim Core1CPath 'путь к исполняемому файлу ядра 1С
Dim ComConnector8 'COM-соединение с 1С8.2 - V82.COMConnector
Dim Ftp 'экземпляр класса для работы с ftp - ChilkatFTP.ChilkatFTP
'------------------------------
Class FtpClass
	Private compFtp
	Private state 'для состояний, true/false
	Private data  'для рассчитанных данных, ассоц. массивы
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
		'проверим не подключены ли к серверу и отключимся, если да
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
		
		'имя подключения связано с подключением через ссылку и проверяется схемой, поэтому проверки получения данных подключения фтп не делаем
		Set ConfigFtpConnection = ConfigXML.selectSingleNode("//*[@idName='" & ConfigFtpUpload.getAttribute("name") & "']")
		compFtp.Hostname = LetParamCompFtp(ConfigFtpConnection,"serverUri")
		compFtp.Port = LetParamCompFtp(ConfigFtpConnection,"serverPort")
		'скорее всего комп со скриптом стоит в локалке за шлюзом, а ftp-сервер смотрит в инет, поэтому по умолчанию пассивный режим,
		'но можно и отключить в конфиг. файле
		compFtp.Passive = IIf(LetParamCompFtp(ConfigFtpConnection,"passiveMode") = "0",false,true)
		compFtp.Username = LetParamCompFtp(ConfigFtpConnection,"login")
		compFtp.Password = LetParamCompFtp(ConfigFtpConnection,"password")
	End Sub
	'-----
	Public Function ChangeRemoteDir(ConfigFtpUpload)    
		'сброс/инициализация предыдущего изменения
		InitState("changeDir") 
		'изменяем рабочую директорию
		If compFtp.ChangeRemoteDir(GetParameterXML(ConfigFtpUpload,"folder")) = 1 Then
			'проверяем изменилась ли она
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
		'сброс/инициализация предыдущего изменения
		InitState("putFile") 
		'перебрасываем файл
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
		'получаем список файлов в XML формате
		FilesListXMLStr = compFtp.GetCurrentDirListing(pattern)
		'парсим и вытягиваем только файлы
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
		'перебрасываем файл
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
	'при отладке можно выводить на экран
	'MsgBox(Message)
	'в рабочем режиме пишем в файл [имя скрипта].log в папке с конфигурационым файлом, или в папку со скриптом при его отсутствии
	If IsEmpty(ConfigFilePath) Then
		FolderLogFilePath = CurrentScriptFolder.Path
	Else
		FolderLogFilePath = fso.GetParentFolderName(ConfigFilePath)
	End If
	WriteToLog Message, FolderLogFilePath
End Sub
'------------------------------
'пишет только при отладке
Sub SendDebug(Message)
    If Debug Then
        SendError "--- " & Message
    End If    
End Sub
'------------------------------
Sub Prepare
	'обработчик файловой системы (общий)
	Set fso = CreateObject("Scripting.FileSystemObject")
	'получаем папку, где лежит файл скрипта
	Set CurrentScriptFolder = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))
	Set Shell = CreateObject("WScript.Shell")
End Sub
'------------------------------
Sub ParseXMLConfigFile(ConfigFilePath)
	Dim FileSchemas,xmlSchema
	'поскольку конфигурационный файл строится на стороне, но надо проверять его по схеме XML (если она есть), чтобы потом не ловить ошибки в коде
	FileSchemas = CurrentScriptFolder.Path & "\config.xsd"
	If fso.FileExists(FileSchemas) Then 
		Set xmlSchema = CreateObject("MSXML2.XMLSchemaCache.6.0")
		xmlSchema.add "", FileSchemas
	End If

	'данные из XML будем забирать по мере необходимости, не распарсивая их сразу
	Set ConfigXML = CreateObject("MSXML2.DOMDocument.6.0")
	If Not IsEmpty(xmlSchema) Then
		Set ConfigXML.schemas = xmlSchema
	End If
	ConfigXML.async="false"

	If Not ConfigXML.load(ConfigFilePath) Then
		SendError "!!!Ошибка валидации конфигурационного файла " & chr(34) & ConfigFilePath & chr(34) & vbCrLf & _
				  "Reason: " & ConfigXML.parseError.reason & vbCrLf & _
				  "Source: " & ConfigXML.parseError.srcText & vbCrLf & _
				  "Line: " & ConfigXML.parseError.Line & vbCrLf & _
				  "Выгрузка прервана."
		Wscript.Quit
	End If
End Sub
'------------------------------
Sub LoadConfig
	'параметры обработки лежат в конфигурационном файле. Он может быть передан в первом параметре при запуске скрипта, либо по 
	'по умолчанию ищем файл конфигурации со стандартным именем config.xml в директории скрипта.
	If WScript.Arguments.Count > 0 Then
		ConfigFilePath = WScript.Arguments(0)
	Else
		ConfigFilePath = CurrentScriptFolder.Path & "\config.xml"
	End If

	If Not fso.FileExists(ConfigFilePath) Then
		SendError "!!!Не найден конфигурационный файл " & chr(34) & ConfigFilePath & chr(34) & ". Выгрузка всех баз прервана."
		WScript.Quit
	Else
		'преобразуем путь к конфигурационному файлу в полный вид
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
	'поиск/создание временной директории для хранения выгрузок
	Dim PathConfig

	If IsEmpty(TempUploadFolder) Then
		'получаем из настроек путь к базе, куда будут делаться выгрузки
		'в случае отсутствия создаем в папке скрипта папку Temp, которую после выгрузки удалим
		'проверяем папку на возможность записи в нее (вдруг по безопасности AD закрыто)
		PathConfig = GetParameterXML(GetNodeXMLRoot("parameters"),"tempUploadFolder")
		If PathConfig = "" Then
			PathConfig = CurrentScriptFolder.Path & "\Temp"
		End If

		'в процессе получения/создания временной директории перехватываем ошибки
		On Error Resume Next
		CreateFolderRecursive PathConfig
		If Not Err.number = 0 Then
			SendError "!!!Ошибка получения временной директории. Выгрузка всех баз прервана. " & Err.Description
			'критическая ошибка - выходим
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
	'для удобства получение корневых веток конф. файла
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
	'получение первого подходящего узла в переданном элементе
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
	'можно не париться и в лоб получить исполняемый файл 1С
	'Core1CPath = "C:\Program Files\1cv82\common\1cestart.exe"
	'но поскольку мы легких путей не ищем, скрипт универсальный, 1С может быть установлена в другой каталог, или версия системы х64, то ищем в реестре откуда он запускается
	'!!!не тестировал на х64 системах с х64 сервером приложений по причине отсутствия лицензионного ключа такого типа
    'в базовых параметрах конфигурации добавил параметр "version", которая указывает используемую версию 1С (может быть 82 или 83)
	On Error Resume Next
	CLSID_1C8Pointer = Shell.RegRead("HKEY_CLASSES_ROOT\V" & GetParameterXML(GetNodeXMLRoot("parameters"),"version") & ".Application\CLSID\")
	Core1CPath = Shell.RegRead("HKEY_CLASSES_ROOT\" & IIf(Is64BitSystem,"Wow6432Node\","") & "CLSID\" & CLSID_1C8Pointer & "\LocalServer32\")
	If Not Err.number = 0 Then
		SendError "!!!Ошибка поиска в реестре пути к исполняемому файлу 1С. " & Err.Description
		SendError "!!!Не найдены все необходимые программы/компоненты для запуска выгрузки. Подготовка к выгрузке всех баз прервана."
		WScript.Quit
	End If
End Function
'------------------------------
Sub ConnectComponent(Var,NameComponent,ComponentDLLPath)
	Dim Component, resultRegister

	On Error Resume Next
	Set Component = CreateObject(NameComponent)
	If Not Err.number = 0 Then
		'не зарегистрирован компонент
		If Err.number = 429 Then
			SendError "!Ошибка " & Err.number & ". Компонента " & chr(34) & NameComponent & chr(34) & " не зарегистрирована в системе."
		ElseIf Err.number = -2147024770 Then
			SendError "!Ошибка " & Err.number & ". Удалили/перенесли файл dll ранее зарегистрированной компоненты " & chr(34) & NameComponent & chr(34)
		Else
			SendError "!Ошибка " & Err.number & ". Не смогли подключить компоненту " & chr(34) & NameComponent & chr(34) & ". " & Err.Description
		End If
		Err.Clear
		'Пробуем зарегить компоненту при наличии dll
		If fso.FileExists(ComponentDLLPath) Then
			'на всякий случай сделаем дерегистрацию компоненты - вдруг там ошибки возникли (удалили dll, а запись в регистре осталась, или записи некорректны)
			Set resultRegister = Shell.Exec("regsvr32.exe /s /u " & chr(34) & ComponentDLLPath & chr(34))
			If Not Err.number = 0 Then
				SendError "!Ошибка при удалении компоненты " & chr(34) & NameComponent & chr(34) &" из файла " & chr(34) & ComponentDLLPath & chr(34) & ". " & Err.Description
				Err.Clear
			Else
				If Not IsEmpty(resultRegister) Then
					'ждем пока компонента удалится
					Do While resultRegister.Status = 0
						WScript.Sleep 100
					Loop
				End If
			End If
			'регистрируем в системе dll-ку в тихом режиме
			Set resultRegister = Shell.Exec("regsvr32.exe /s " & chr(34) & ComponentDLLPath & chr(34))
			If Not Err.number = 0 Then
				SendError "!Ошибка при регистрации компоненты " & chr(34) & NameComponent & chr(34) & " из файла " & chr(34) & ComponentDLLPath & chr(34) & ". " & Err.Description
			Else 
				If Not IsEmpty(resultRegister) Then
					'ждем пока зарегистрируется компонента
					Do While resultRegister.Status = 0
						WScript.Sleep 100
					Loop
				End If
				SendError "Подключили в тихом режиме компоненту " & chr(34) & NameComponent & chr(34) & " из файла " & chr(34) & ComponentDLLPath & chr(34) & ". " & Err.Description
				'пробуем опять получить компонент. Если ошибка, то в лог для разборок в ручном режиме
				Set Component = CreateObject(NameComponent)
				If Not Err.number = 0 Then
					SendError "!Ошибка подключения компоненты " & chr(34) & NameComponent & chr(34) & " после установки. " & Err.Description
					SendError "!Проверьте правильность запуска скрипта. Cм. файл readme.txt"
				Else
					Set Var = Component
				End If
			End If
		Else
			SendError "!Не найден файл " & chr(34) & ComponentDLLPath & chr(34) & " для регистрации компоненты " & chr(34) & NameComponent & chr(34)
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
	'COM-коннектор лежит в папке с исполняемым файлом 1С
	'подключаем тогда, когда есть активные на выгрузку клиент-серверные базы
	If GetPresentActivityClientServerBases() Then
		ConnectComponent ComConnector8, "V" & GetParameterXML(GetNodeXMLRoot("parameters"),"version") & ".COMConnector", fso.GetParentFolderName(Core1CPath) & "\comcntr.dll"
		If IsEmpty(ComConnector8) Then
			SendError "!!Выгрузка баз с сервера приложений будет пропущена, т.к. не подключен COM-коннектор 1С " & GetParameterXML(GetNodeXMLRoot("parameters"),"version") & "."
		End If
	End If
End Function
'------------------------------
Function ConnectFtpComponent
	Dim ComponentDLLPath, Component
	'в качестве компоненты для работы с ftp используем ChilkatFtp. Она freeware, легко интегрируется в систему и достаточна для простых операций.
	'с сайта производителя http://www.chilkatsoft.com/downloads.asp она уже убрана, поэтому dll и help-файл с командами лежит в папке ftp репозитория
	'ее можно переложить в другое место (к системным библиотекам) - путь к dll указываем в конфиг-файле в ветке root/parameters/ftpComponentPathDLL
	'P.S. Встроенный в систему ftp.exe не впечатлил своими возможностями и методами работы с ним
	'подключение компоненты осуществляем только если есть настройки фтп-подключений в конфигурационном файле, а также базы с активными настройками выгрузки на фтп
	If Not (GetNodeXMLRoot("ftpConnections") Is Nothing) And GetPresentActivityFtpUploadBases() Then
		ComponentDLLPath = GetParameterXML(GetNodeXMLRoot("parameters"),"ftpComponentPathDLL")
		ComponentDLLPath = IIf(ComponentDLLPath = "",CurrentScriptFolder.Path & "\ftp\ChilkatFTP.dll",ComponentDLLPath)
		ConnectComponent Component, "ChilkatFTP.ChilkatFTP", ComponentDLLPath
		If IsEmpty(Component) Then
			SendError "!!Переброс файлов на ftp-сервер не будет выполнен. Выгрузка баз будет осуществлена во временную директорию."
		Else
			Set Ftp = New FtpClass
			Set Ftp.Component = Component
		End If
	End If
End Function
'------------------------------
'добавляет данные в фиксированный массив, увеличивая его размер
Function AddInFixedArray(VarArray,VarData)
    Dim CurrentSizeVarArray
    
    CurrentSizeVarArray = UBound(VarArray,1)
	'если нулевой элемент не пустой - тогда увеличиваем размер
    If Not IsEmpty(VarArray(0)) Then
		'увеличиваем размер динамического массива на 1
        CurrentSizeVarArray = CurrentSizeVarArray + 1
		ReDim Preserve VarArray(CurrentSizeVarArray)
	End If
    Set VarArray(CurrentSizeVarArray) = VarData
End Function
'------------------------------
Function FindBaseOnClientServer(dicUploadData)
	Dim Cluster
	'рабочий процесс
    Dim WorkingProcess
	'список всех рабочих процессов
    ReDim ConnectWorkingProcesses(0)
    'текущий рабочий процесс
    Dim ConnectWorkingProcess
    'база, которую ищем в процессах
	Dim TempBase

	'подсоединимся к агенту кластера 1С
	On Error Resume Next
	
	If Err.number = 0 Then
		On Error Goto 0
		
        Set Cluster = dicUploadData.Item("COMAgent1C").GetClusters()(0)
        'считаем что администраторы кластера не введены в систему - используем пустую строку
		'иначе нужно дописать параметры администратора кластера в конфиграционный файл и тут проверить
		dicUploadData.Item("COMAgent1C").Authenticate Cluster, "", ""
		
		For Each WorkingProcess in dicUploadData.Item("COMAgent1C").GetWorkingProcesses(Cluster)
            'заметка - не путаем понятия из 1С: рабочий процесс и соединение с рабочим процессом
            'только соединение с рабочим процессом имеет методы() для работы с ним
			Set ConnectWorkingProcess = ComConnector8.ConnectWorkingProcess("tcp://"+dicUploadData.Item("BaseConnection").Item("serverName")+":"+CStr(WorkingProcess.MainPort))		
            'объект ConnectWorkingProcess будет нам нужен для дисконнекта пользователей
			'сразу залогинимся. здесь нужен доступ уровня администратор
            ConnectWorkingProcess.AddAuthentication dicUploadData.Item("BaseConnection").Item("login"), dicUploadData.Item("BaseConnection").Item("password")
	    
            AddInFixedArray ConnectWorkingProcesses, ConnectWorkingProcess
		Next
        dicUploadData.Add "ConnectWorkingProcessesClientServer", ConnectWorkingProcesses

        'перебираем базы в одном из рабочих процессов, и ищем нужную нам
        For Each TempBase in ConnectWorkingProcesses(0).GetInfoBases()
			If TempBase.name = dicUploadData.Item("BaseConnection").Item("baseName") Then
				'нашли инфомационную базу, добавляем в общие данные
				dicUploadData.Add "BaseClientServer", TempBase
				Exit For
			End If
		Next
	Else
		SendError "!!Ошибка подключения к серверу приложений для базы " & FormatBaseNameLog(dicUploadData) & ". Выгрузка выбранной базы прервана. Описание ошибки: " & Err.Description
		Err.Clear
	End If
	FindBaseOnClientServer = dicUploadData.Exists("BaseClientServer")
End Function
'------------------------------
Function DisconnectUsers(dicUploadData)
	Dim CurrentConnections, CurrentConnection
    Dim ConnectWorkingProcess

	DisconnectUsers = True
	'удаляем все активные соединения, если они не в монопольном режиме (в монопольном работает конфигуратор и ком-соединение)...
	On Error Resume Next
	
    For Each ConnectWorkingProcess in dicUploadData.Item("ConnectWorkingProcessesClientServer")
        CurrentConnections = ConnectWorkingProcess.GetInfoBaseConnections(dicUploadData.Item("BaseClientServer"))
	    If Not Err.number = 0 Then
		    SendError "!!Ошибка получения текущих соединений базы " & FormatBaseNameLog(dicUploadData) & ". Проверьте параметры аутентификации. " & Err.Description
		    Err.Clear
		    DisconnectUsers = False
	    Else
		    For Each CurrentConnection in CurrentConnections
			    'проверяем только рабочие доступы. Конфигурирование, консоль и текущее Com-соединение не трогаем
			    If CurrentConnection.AppID = "1CV8" Then
				    If CurrentConnection.IBConnMode = 0 Then
					    ConnectWorkingProcess.Disconnect(CurrentConnection)
				    Else
					    'есть запущенное предприятие в монопольном режиме - наверное проведение, пропускаем выгрузку
                        SendError "!Открыты монопольные соединения, выгрузка пропущена"
					    DisconnectUsers = False
				    End If
			    ElseIf CurrentConnection.AppID = "Designer" Then
				    'открыт конфигуратор - пропускаем выгрузку (хотя можно и закрыть его! Это чтобы изменения при разработке вдруг не потерялись)
                    SendError "!Открыт конфигуратор, выгрузка пропущена"
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
	'если указана папка выгрузки в настройках базы - то проверяем ее наличие/доступ и выгружаем в нее, иначе во временную папку
	GetFolderUploadPath = GetParameterXML(dicUploadData.Item("BaseConfig"),"baseUploadFolder")
	IF Not GetFolderUploadPath = "" Then
		CreateFolderRecursive GetFolderUploadPath
		If Not Err.number = 0 Then
			'если возникли сложности с доступом к папке выгрузки - пишем ошибку и продолжаем выгрузку во временную папку
			SendError "!Ошибка доступа/создания папки " & chr(34) & GetFolderUploadPath & chr(34) & " для базы " & FormatBaseNameLog(dicUploadData) & _
					  ". Продолжаем выгрузку во временную папку. Описание ошибки: " & Err.Description
			Err.Clear
			GetFolderUploadPath = GetTempUploadFolderPath
		End If
	Else
		GetFolderUploadPath = GetTempUploadFolderPath
	End If
End Function
'------------------------------
Function GetFileUploadPath(dicUploadData)
	'вынес в отдельную процедуру, чтобы можно было манипулировать именем файла и стандартизировать его создание
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
	'сюда приходят уже все необходимые данные для подключения и переброски
	FtpUpload = false
	If Ftp.Connect(dicUploadData.Item("ConfigFtpUpload")) Then
		If Ftp.ChangeRemoteDir(dicUploadData.Item("ConfigFtpUpload")) Then
			If Ftp.PutFile(dicUploadData.Item("FileUploadPath"),fso.GetFileName(dicUploadData.Item("FileUploadPath"))) Then
				SendError "Перебросили выгрузку " & chr(34) & dicUploadData.Item("FileUploadPath") & chr(34) & " на ftp-сервер с конфиг. именем " & chr(34) & dicUploadData.Item("ConfigFtpUpload").getAttribute("name") & chr(34) & "."
				FtpUpload = true
			Else
				SendError "!!Не смогли передать файл выгрузки на ftp-сервер " & chr(34) & Ftp.GetParamFtp("serverUri") & chr(34) & ". Переброска выгрузки "& chr(34) & _
						  dicUploadData.Item("FileUploadPath") & chr(34) & " на ftp-сервер прервана. " & vbCrLf & Ftp.GetState("lastError")
			End If
		Else
			SendError "!!Не смогли изменить рабочую директорию на ftp-сервере " & chr(34) & Ftp.GetParamFtp("serverUri") & chr(34) & ". Переброска выгрузки " & chr(34) & _
				  dicUploadData.Item("FileUploadPath") & chr(34) & " на ftp-сервер прервана. " & vbCrLf & Ftp.GetState("lastError")
		End If
	Else
		SendError "!!Ошибка подключения к ftp-серверу с конфиг. именем " & chr(34) & dicUploadData.Item("ConfigFtpUpload").getAttribute("name") & chr(34) & ". Переброска выгрузки " & chr(34) & _
				  dicUploadData.Item("FileUploadPath") & chr(34) & " на ftp-сервер прервана. " & vbCrLf & Ftp.GetState("lastError")
	End If
	Ftp.Disconnect()
End Function
'------------------------------
Function UploadBaseFtp(dicUploadData)
	UploadBaseFtp = False
	If Not IsEmpty(Ftp) Then
		UploadBaseFtp = FtpUpload(dicUploadData)
		If UploadBaseFtp Then
			'нужно ли удалять файл из временной директории после переброски. По умолчанию удаляем.
			If Not dicUploadData.Item("ConfigFtpUpload").getAttribute("deleteAfterUpload") = "0" Then
				fso.DeleteFile(dicUploadData.Item("FileUploadPath"))
			End If
		End If
	Else
		SendError "!!Ftp компонента не подключена. Переброска выгрузки " & chr(34) & dicUploadData.Item("FileUploadPath") & chr(34) & " на ftp-сервер прервана."
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

	'папка выгрузки пригодится при подчистке файлов выгрузок
    'для контроля выгрузки создаем файл вывода информации по выгрузке (для вставки в журнал информации в случае ошибки)
    'а также файл дампа для контроля правильности выгрузки (там флаги 0,1, 101)
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
		'анализируем процесс выгрузки, разбирая файлы ответа и дампа
        DumpResult = ReadTextFileUTF8(dicUploadData.Item("DumpResultPath"))
        If DumpResult = "0" Then
            dicUploadData.Item("ResultUploadBaseDT") = True
		    SendError "Выгрузили базу " & FormatBaseNameLog(dicUploadData) & " в " & chr(34) & dicUploadData.Item("FileUploadPath") & chr(34)
        Else
           Set FileOutResult = fso.OpenTextFile(dicUploadData.Item("OutResultPath"), 1, False) 
           SendError Trim("!!Ошибка при выгрузке базы " & FormatBaseNameLog(dicUploadData) & ": " & FileOutResult.ReadAll)
           FileOutResult.Close
        End If
                
        WScript.Sleep 1000
        'зачищаем файлы дампа
        fso.DeleteFile dicUploadData.Item("DumpResultPath")
        fso.DeleteFile dicUploadData.Item("OutResultPath")
	End If	
End Sub
'------------------------------
Sub UploadBaseClientServer(dicUploadData)
	ConnectAgent1C dicUploadData
    
    If FindBaseOnClientServer(dicUploadData) Then
		'ставим блокировку на базу
        BlockBase dicUploadData, true
        'выгоняем пользователей с базы
		If DisconnectUsers(dicUploadData) Then
			'если всех выгнали - производим выгрузку. Блокировку базы при выгрузке не ставил, т.к. обычно не успевают заходить, базы выгружаются ночью и при выгрузке база блокируется.
			UploadBaseClientServer2 dicUploadData
			If Not dicUploadData.Item("ResultUploadBaseDT") Then
				SendError "!!Не смогли выгрузить базу " & FormatBaseNameLog(dicUploadData) & " через запуск командной строки с параметрами."
			End If
		Else
			SendError "!!Не смогли отключить всех пользователей для базы " & FormatBaseNameLog(dicUploadData) & ". Выгрузка выбранной базы прервана."
		End If
        'снимаем блокировку с базы
        BlockBase dicUploadData, false
	Else
		SendError "!!Не найдена на сервере база " & FormatBaseNameLog(dicUploadData) & ". Выгрузка выбранной базы прервана."
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
'закрываем все открытые подключения к 1С - агент и соединения с рабочими процессами
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
	'проверяем префикс, чтобы не зацепить чужие файлы. Расширение файла может быть любым.
	lenPrefixName = Len(dicUploadData.Item("fileNamePrefixUpload"))
	If Left(fso.GetFileName(File),lenPrefixName) = dicUploadData.Item("fileNamePrefixUpload") Then
		'обрезаем префикс и парсим дату файла
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
	'разница между текущей датой и началом исследуемого месяца
	RestructuringFilesMonthCheck = (depthInMonth = 0) OR (DateDiff("m","01/" & tempMonth & "/" & tempYear,date) < depthInMonth)
End Function
'------------------------------
Sub RestructuringFilesSelect(dicUploadData)
	Dim PatternElement, rangeInDay
	Dim ArrYears, tempYear, ArrMonth, tempMonth, ArrDay, tempDay, tempDate, lastDate, i, FileTempPath

	'получаем файлы, которые оставляем
	dicUploadData.Add "RestructuringFilesLeave", CreateObject("Scripting.Dictionary")
	For Each PatternElement In dicUploadData.Item("restructuringPattern").childNodes
		ArrYears = dicUploadData.Item("RestructuringFilesTree").Keys
		For Each tempYear In ArrYears
			ArrMonth = dicUploadData.Item("RestructuringFilesTree").Item(tempYear).Keys
			For Each tempMonth In ArrMonth
				'проверяем, подходит ли нам текущий месяц по глубине
				If RestructuringFilesMonthCheck(tempYear,tempMonth,CInt(GetParameterXML(PatternElement,"depthInMonth"))) Then
					rangeInDay = CInt(GetParameterXML(PatternElement,"rangeInDay"))
					'дни внутри месяца сканируем в обратном порядке, с конца месяца
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
	'получаем файлы, которые удаляем
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
		'выбираем все файлы с папки выгрузки, и заносим в структуру-дерево с разбивкой по годам/месяцам
		For Each CurrentFile In Ftp.GetListFiles("*.*") 
			RestructuringFilesAdd dicUploadData, CurrentFile
		Next
		'формируем список файлов, которые удаляем
		RestructuringFilesSelect dicUploadData
		'удаляем файлы
		For Each CurrentFile In dicUploadData.Item("RestructuringFilesDelete").Items
			Ftp.DeleteFile CurrentFile
		Next
	End If
End Sub
'------------------------------
Sub RestructuringLocalDT(dicUploadData)
	Dim CurrentFile
	'выбираем все файлы с локальной папки выгрузки, и заносим в структуру-дерево с разбивкой по годам/месяцам
	For Each CurrentFile In fso.GetFolder(dicUploadData.Item("FolderUploadPath")).Files 
		RestructuringFilesAdd dicUploadData, CurrentFile
	Next
	'формируем список файлов, которые удаляем
	RestructuringFilesSelect dicUploadData
	'удаляем файлы
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
		'получаем паттерн реструктурирования файлов
		dicUploadData.Add "restructuringPattern", ConfigXML.selectSingleNode("//*[@idName='" & RestructuringLink.getAttribute("patternName") & "']") 
		'получаем список файлов, строим на их основании список тех, которые оставить, а неподходящие удаляем
		If dicUploadData.Item("ResultUploadBaseDT") Then		
			If dicUploadData.Item("ResultUploadBaseFtp") = True Then
				'чистим папку на фтп-сервере
				RestructuringFtpDT dicUploadData
				SendError "Провели реструктуризацию папки " & chr(34) & GetParameterXML(dicUploadData.Item("ConfigFtpUpload"),"folder") & chr(34) & _
						  " ftp-сервера с конфиг. именем " & chr(34) & dicUploadData.Item("ConfigFtpUpload").getAttribute("name") & chr(34)
			Else
				'локальная папка
				RestructuringLocalDT dicUploadData
				SendError "Провели реструктуризацию локальной папки " & chr(34) & dicUploadData.Item("FolderUploadPath") & chr(34)
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
	'соберем все параметры подключения к базе в удобночитаемом виде
	Set dicBaseConnection = CreateObject("Scripting.Dictionary")
	dicUploadData.Add "BaseConnection", dicBaseConnection
	dicBaseConnection.Add "login", GetParameterXML(BaseConnection,"login")
	dicBaseConnection.Add "password", GetParameterXML(BaseConnection,"password")
	'проверяем тип подключения базы 1С и производим выгрузку
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

	'определяем режим отладки
    Debug = (GetParameterXML(GetNodeXMLRoot("parameters"),"debug") = "1")
    'перебираем ветки с базами из конфигурационного файла и по очереди их выгружаем
	For Each BaseConfig in GetNodeXMLRoot("bases").childNodes
		'создаем переменную для сбора/передачи данных текущей базы между процедурами
		SendDebug "start create dictionary"
        Set dicUploadData = CreateObject("Scripting.Dictionary")
		SendDebug "end create dictionary"
        dicUploadData.Add "BaseConfig", BaseConfig
        
		'базы можно убирать из цикла выгрузки атрибутом <activity> = 0, все другие значения и отсутствие атрибута игнорируется
		SendDebug "start check activity base"
        If Not BaseConfig.getAttribute("activity") = "0" Then
			SendDebug "end check activity base"
            'производим выгрузку базы
			UploadBase dicUploadData            
			'перебрасываем на фтп-сервер при необходимости
			TransferToFtp dicUploadData
			'проводим подчистку папки с выгрузками
			RestructuringBase dicUploadData
		End If
        
        SendDebug "start clear dictionary"
        Set dicUploadData = nothing
		SendDebug "end clear dictionary"
	Next
End Sub
'------------------------------
Sub Run
    Prepare	'подготовка общих переменных
    LoadConfig 'загрузка конфигурационного файла
    FindCore1C 'поиск ядра 1С
    ConnectComConnector1C8 'подключение компоненты COM-коннектора к 1С 8
    ConnectFtpComponent 'подключение компоненты для работы с ftp (при необходимости)
    UploadBases

    SendError "--------------------------------------"
End Sub
'------------------------------
Run