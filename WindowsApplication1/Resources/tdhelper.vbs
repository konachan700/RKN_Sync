Option Explicit

' Для вызова справки запустите данный скрипт без параметров

' Определяет директорию для хранения log-файлов
const LOG_PATH = "./logs"


' ===== History
' 24.09.2009, версия 1.0, ООО "Цифровые технологии"
'
' Реализовано:
' - выполнение следующих операций над подходящими файлами в заданном каталоге:
'   - установка ЭЦП, в т.ч. УЭЦП, а также просто со штампами времени
'   - проверка ЭЦП, в т.ч. заверяющих
'   - снятие ЭЦП
'   - шифрование (с проверкой на недействительность сертификатов получателей)
'   - расшифрование
' - результат операций записывается в файл журнала


' ===== Constants

' enum PROFILESTORETYPE
Const REGISTRY_STORE = 0
Const XML_STORE = 1 ' since TD 3.3

' enum DATA_TYPE
Const DT_EMPTY_SIGNED_DATA = -2 ' since TD 4.2
Const DT_AUTO_DETECT = -1
Const DT_PLAIN_DATA = 0
Const DT_SIGNED_DATA = 2
Const DT_ENVELOPED_DATA = 3

' enum FORMAT 
Const UNKNOWN_TYPE = -1 ' since TD 4.5
Const BASE64_TYPE = 0 ' is equal to PROFILEEXITFORMAT::BASE64
Const DER_TYPE = 1 ' is equal to PROFILEEXITFORMAT::DER
Const XML_TYPE = 2 ' since 3.3
Const HEX_TYPE = 3 ' since 4.3
Const BINARY_TYPE = 1 ' is equal to DER_TYPE ' since TD 4.4
Const UNICODE_STRING_TYPE = 4 ' since TD 4.4

' enum POLICY_TYPE
Const POLICY_TYPE_NONE = 0
Const POLICY_TYPE_SIGNATURE = 1
Const POLICY_TYPE_ENCRYPT = 2

' enum SIGNATURE_TYPE
Const SIGNATURE_TYPE_UNKNOWN = 0            ' // неизвестный тип подписи
Const SIGNATURE_TYPE_BASIC = 1              ' // обычная подпись (не путать с CAdES-BES)
' // CAdES types
' //SIGNATURE_TYPE_CADES_BES = 2         // CAdES Basic Electronic Signature
' //SIGNATURE_TYPE_CADES_EPES = 4        // CAdES Explicit Policy Electronic Signatures
' //SIGNATURE_TYPE_CADES_T = 8           // CAdES with Time
' //SIGNATURE_TYPE_CADES_C = 16          // CAdES with Complete validation data references
' //SIGNATURE_TYPE_CADES_X_LONG = 32     // Extended validation data: Long validation data
' //SIGNATURE_TYPE_CADES_X_TYPE_1 = 64   // Extended validation data: Type 1
' //SIGNATURE_TYPE_CADES_X_TYPE_2 = 128  // Extended validation data: Type 2
Const SIGNATURE_TYPE_CADES_X_LONG_TYPE_1 = 96 '//= SIGNATURE_TYPE_CADES_X_LONG + SIGNATURE_TYPE_CADES_X_TYPE_1

' enum VERIFYFLAG
Const VF_CERT_AND_SIGN_OLD = 0
Const VF_SIGN_ONLY = 1
Const VF_CERT_ONLY = 2 ' // only using with VF_SIGN_ONLY
Const VF_CERT_AND_SIGN = 3
Const VF_TSP_ONLY = 4 ' // only using with VF_SIGN_ONLY
Const VF_SIGN_AND_TSP = 5
'Const VF_SIGN_AND_CERT_AND_TSP = 6
Const VF_ALL_POSSIBLE = -1

' enum VERIFYSTATUS
Const VS_CORRECT = 1                             ' // Проверка прошла успешно
Const VS_UNSUFFICIENT_INFO = 2                   ' // Недостаточно данных (возникает, когда нет СОС для проверки сертификата вышестоящего УЦ)
Const VS_UNCORRECT = 3                           ' // ЭЦП сертификата, СОС или СДС не корректна
'
Const VS_INVALID_CERTIFICATE_BLOB = 4            ' // Не создан контекст сертификата
Const VS_CERTIFICATE_TIME_EXPIRIED = 5           ' // Истёк срок действия сертификата
Const VS_CERTIFICATE_NO_CHAIN = 6                ' // Невозможно построить цепочку (путь сертификации) для сертификата
Const VS_CRL_UPDATING_ERROR = 7                  ' // Произошла ошибка при обновлении СОС
Const VS_LOCAL_CRL_NOT_FOUND = 8                 ' // Не найден СОС в локальном хранилище
Const VS_CRL_TIME_EXPIRIED = 9                   ' // СОС найден, однако он нуждается в обновлении
Const VS_CERTIFICATE_IN_CRL = 10                 ' // Сертификат содержится в СОС
Const VS_CERTIFICATE_IN_LOCAL_CRL = 11           ' // Сертификат содержится в СОС, обновление СОС не было затребовано
Const VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL = 12   ' // Сертификат действителен, но обновление СОС не было затребовано
Const VS_CERTIFICATE_USING_RESTRICTED = 13       ' // Сертификат запрещен для использования (настройками, политикой и др.)
Const VS_NOT_APPLICABLE_FOR_SPECIFIED_USAGE = 13 ' // Не применим для указанной области использования (для СДС)
Const VS_CERTIFICATE_RESTRICTED_BY_LICENSE = 14  ' // Отсутствует лицензия на использование данного сертификата
Const VS_REVOCATION_STATUS_UNKNOWN = 15          ' // RP возвращает статус UNKNOWN (лицензия или ещё что-другое)
Const VS_REVOCATION_OCSP_ERROR = 16              ' // По каким-либо причинам не удалось проверить статус сертификата по OCSP
Const VS_CADES_ATTRIBUTES_NOT_VERIFIED = 17      ' // ЭЦП корректна, но дополнительные атрибуты УЭЦП не учитывались при проверке
Const VS_CHAIN_UNCORRECT_BY_SPECIFIED_CTLS = 18  ' // Доверенность пути сертификации не гарантируется (ни одним из) СДС
Const VS_CTL_IS_NOT_SIGNED = 19                  ' // СДС не подписан (для сертификата - если включена проверка по СДС, для СДС - при наличии флага требовать подпись СДС)
Const VS_CTL_TIME_EXPIRIED = 20                  ' // Время действия СДС истекло (или еще не началось)

' enum WIZARD_TYPE
Const SIGN_WIZARD_TYPE = 1
Const ADD_SIGN_WIZARD_TYPE = 2
Const COSIGN_WIZARD_TYPE = 4
Const CONNECT_KEYCARRIER_WIZARD_TYPE = 8 ' since TD 3.2
Const ENCRYPT_WIZARD_TYPE = 64
Const DECRYPT_WIZARD_TYPE = 1024
Const VERIFY_SIGNATURE_WIZARD_TYPE = 2048
Const DECRYPT_VERIFY_SIGNATURE_WIZARD_TYPE = 4096
Const DROP_SIGNATURE_WIZARD_TYPE = 8192
Const VIEW_DOCUMENT_WIZARD_TYPE = 16384
Const OPEN_MANAGERS_WIZARD_TYPE = 32768
'//COSIGN_WIZARD_TYPE_WITH_FILES = 32768
Const ADD_SIGN_WIZARD_TYPE_WITH_FILES = 65536


' self constants
Const STATUS_UNKNOWN = -1
Const STATUS_OK = 0
Const STATUS_WARNING = 1
Const STATUS_BAD = 2
Const STATUS_INVALID = 3


' ===== global variables
Dim g_iOperationMode : g_iOperationMode = ""
Dim g_sSignExt : g_sSignExt = ""
Dim g_sEncrExt : g_sEncrExt = ""

Dim g_oProfile : Set g_oProfile = Nothing

Dim g_oPKCS7Message1 : Set g_oPKCS7Message1 = CreateObject("DigtCrypto.PKCS7Message")
Dim g_oPKCS7Message2 : Set g_oPKCS7Message2 = Nothing
Dim g_oPKCS7Message : Set g_oPKCS7Message = g_oPKCS7Message1
Dim g_oPKCS7MessageShared : Set g_oPKCS7MessageShared = g_oPKCS7Message1

Dim oFSOShared : Set oFSOShared = CreateObject("Scripting.FileSystemObject")

Dim oLogFile : Set oLogFile = Nothing



' ===== helpers
Function IsNothing( ByRef oObject )
	if Not IsObject(oObject) then
		IsNothing = True
	elseif TypeName(oObject) = "Nothing" then
		IsNothing = True
	else
		IsNothing = False
	end if
End function

Function TruncatePath( ByVal sFullname )
	TruncatePath = sFullname

	Dim iPos : iPos = InStrRev( sFullname, "\" )
	if IsNumeric(iPos) then
		if iPos > 0 then
			TruncatePath = Right( sFullname, Len(sFullname) - iPos )
		end if
	end if
End Function

Function TruncateFileExtension( ByVal sFullname )
	TruncateFileExtension = sFullname

	Dim sFilename : sFilename = TruncatePath(sFullname)
	Dim iPos : iPos = InStrRev( sFilename, "." )
	if IsNumeric(iPos) then
		if iPos > 0 then
			TruncateFileExtension = Left( sFullname, Len(sFullname) - Len(sFilename) + iPos - 1 )
		end if
	end if
End Function

Function InitLog()
	Dim curDate : curDate = Date()
	Dim curYear : curYear = Right(CStr(Year(curDate)), 2)
	Dim curMonth : curMonth = CStr(Month(curDate)) : curMonth = String(2-len(curMonth), "0") & curMonth
	Dim curDay : curDay = CStr(Day(curDate)) : curDay = String(2-len(curDay), "0") & curDay

	Dim sLogFileName : sLogFileName = oFSOShared.BuildPath(LOG_PATH, curYear & curMonth & curDay & ".log")

	if not oFSOShared.FolderExists( LOG_PATH ) then
		oFSOShared.CreateFolder( LOG_PATH )
	end if
	Set oLogFile = oFSOShared.OpenTextFile( sLogFileName, 8, true )


	Dim sParams : sParams = ""
	Dim i, iCount : iCount = WScript.Arguments.Count
	For i = 0 to iCount - 1
		sParams = sParams & """" & WScript.Arguments.Item(i) & """ "
	Next

	oLogFile.WriteLine ""
	LogMsg( "Скрипт: """ & WScript.ScriptFullName & """" )
	LogMsg( "Запущен с параметрами: " & sParams )
	LogMsg( "----------------------------------------------------------------------------" )
End Function

Function LogMsg( ByVal sMessage )
	oLogFile.WriteLine( CStr(Time()) & " " & Replace(sMessage, vbCrLf, " / ") )

	if InStr( WScript.FullName, "cscript.exe" ) > 0 then
		WScript.Echo sMessage
	end if
End Function

Function PrintError( ByVal sMessage, ByVal iExitCode ) ' if iExitCode == -1, then iExitCode set to Err.Number
	if -1 = iExitCode then
		iExitCode = Err.Number
	end if

	Dim sRes : sRes = "Error number: " & CStr(Err.Number) & ", Description: " & CStr(Err.Description)

	if 0 = iExitCode then
		sRes = sMessage
	elseif Len(sRes) > 0 then
		sRes = sMessage & vbCrLf & sRes
	end if

	LogMsg sRes
	'WScript.Echo sRes

	WScript.Quit iExitCode
End Function

Function PrintUsage( ByVal iExitCode )
	Dim sUsage : sUsage = _
		"Описание использования скрипта: " & vbCrLf & _
		WScript.ScriptName & " <вх_каталог> <вых_каталог> <имя_настройки_КриптоАРМ> <операция>" & vbCrLf & _
		vbCrLf & _
		"параметр <операция> может принимать соледующие значения:" & vbCrLf & _
		"- S[ign] - подписать" & vbCrLf & _
		"- V[erifySignature] - проверить подпись" & vbCrLf & _
		"- T[akeOffSignature] - снять подпись" & vbCrLf & _
		"- E[ncrypt] - зашифровать" & vbCrLf & _
		"- D[ecrypt] - расшифровать"

	WScript.Echo sUsage
	WScript.Quit iExitCode
End Function

Function CachePinCode() ' implemented caching engine is applicable for CryptoPro CSP (only?)
	' caching of pin-code
	if IsNothing(g_oPKCS7Message2) then
'		g_oPKCS7Message.Profile = g_oProfile
'		g_oPKCS7Message.Import DT_PLAIN_DATA, "123"
'		g_oPKCS7Message.Sign
'		g_oPKCS7Message.Export DT_SIGNED_DATA, BASE64_TYPE

		Set g_oPKCS7Message2 = CreateObject("DigtCrypto.PKCS7Message")
		g_oPKCS7Message2.Profile = g_oPKCS7Message.Profile
		Set g_oPKCS7Message = g_oPKCS7Message2
	end if
End Function

Function ParseCryptoOperation( ByVal sOperation )

	if 0 = StrComp( Left(sOperation, 1), "s", vbTextCompare ) then
		ParseCryptoOperation = SIGN_WIZARD_TYPE
	elseif 0 = StrComp( Left(sOperation, 1), "t", vbTextCompare ) then
		ParseCryptoOperation = DROP_SIGNATURE_WIZARD_TYPE
	elseif 0 = StrComp( Left(sOperation, 1), "v", vbTextCompare ) then
		ParseCryptoOperation = VERIFY_SIGNATURE_WIZARD_TYPE
	elseif 0 = StrComp( Left(sOperation, 1), "e", vbTextCompare ) then
		ParseCryptoOperation = ENCRYPT_WIZARD_TYPE
	elseif 0 = StrComp( Left(sOperation, 1), "d", vbTextCompare ) then
		ParseCryptoOperation = DECRYPT_WIZARD_TYPE
	else
		PrintError "Указана неподдерживаемая операция: """ & sOperation & """", -1
	end if
End Function

Function LoadProfileByIdOrName( ByVal sProfileName )
	Dim oProfileStore : Set oProfileStore = CreateObject("DigtCrypto.ProfileStore")
	oProfileStore.Open REGISTRY_STORE, ""

	Dim oProfiles : Set oProfiles = oProfileStore.Store
	Dim oResProfile : Set oResProfile = oProfiles.Profile( sProfileName )

	if IsNothing(oResProfile) then
		Dim iCount : iCount = oProfiles.Count

		Dim oProfile, i
		For i = 0 to iCount - 1
			Set oProfile = oProfiles.Item(i)
			if oProfile.Name = sProfileName then
				Set oResProfile = oProfile
				exit for
			end if
		Next
	end if

	if not IsNothing(oResProfile) then
		if Len(oResProfile.TSPProfileID) > 0 then ' in TD 4.5.0 it's needs to be set of TSP&OCSP profiles manually
			On Error Resume Next
			Dim oTSPProfile : Set oTSPProfile = oResProfile.TSPProfile ' in TD 4.5.0 execption thowing if it's not set
			if Err.Number <> 0 then
				Set oTSPProfile = Nothing
			end if
			On Error Goto 0

			if IsNothing(oTSPProfile) then
				oResProfile.TSPProfile = oProfileStore.TSPProfileStore.Profile(oResProfile.TSPProfileID)
			end if
		end if
	end if

	Set LoadProfileByIdOrName = oResProfile
End Function

Function VerifyProfile( ByRef oProfile, ByVal iOperation )
	if ENCRYPT_WIZARD_TYPE = iOperation then
		' checking of recipient certificates
		Dim oRecipients : Set oRecipients = oProfile.Recipients
		Dim oCert, i, c : c = oRecipients.Count

		Dim iCurStatus, iResStatus : iResStatus = STATUS_UNKNOWN
		For i = 0 to c - 1
			Set oCert = oRecipients.Item(i)
			iCurStatus = VerifyCertificate( oCert, POLICY_TYPE_ENCRYPT )
			iCurStatus = SignStatusToSolidStatus( iCurStatus )

			if iResStatus < iCurStatus then
				iResStatus = iCurStatus
			end if
		Next

		if iResStatus >= STATUS_BAD then
			PrintError "Нет доверия к одному или нескольким сертификатам получателей шифрованного сообщения", -1
		end if
	end if
End Function



' ===== functions

' ---
Function SignFile( ByVal sInputFilename, ByVal sOutputFilename )
	g_oPKCS7Message.Load DT_PLAIN_DATA, sInputFilename
	g_oPKCS7Message.Sign
	g_oPKCS7Message.Save DT_SIGNED_DATA, g_oProfile.SignExitFormat, sOutputFilename

	LogMsg "Файл """ & sInputFilename & """ подписан и сохранен в """ & sOutputFilename & """"

	CachePinCode
End Function

' ---
Function TakeOffSignature( ByVal sInputFilename, ByVal sOutputFilename )
	On Error Resume Next

	Dim bAttached : bAttached = g_oPKCS7Message.Load( DT_SIGNED_DATA, sInputFilename )

	if Err.Number <> 0 then
		LogMsg "Файл """ & sInputFilename & """ не является ЭЦП!"
	elseif not bAttached then
		LogMsg "Файл ЭЦП """ & sInputFilename & """ не содержит подписанного документа!"
	else
		g_oPKCS7Message.Save DT_PLAIN_DATA, DER_TYPE, sOutputFilename

		LogMsg "Из файла ЭЦП """ & sInputFilename & """ исходный документ извлечен в """ & sOutputFilename & """"
	end if

	On Error Goto 0
End Function

' ---
Function SignStatusToSolidStatus( ByVal iSignStatus )
	Select Case iSignStatus

	Case VS_CORRECT SignStatusToSolidStatus = STATUS_OK
	Case VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL SignStatusToSolidStatus = STATUS_OK

	Case VS_UNSUFFICIENT_INFO SignStatusToSolidStatus = STATUS_WARNING
	Case VS_CRL_UPDATING_ERROR SignStatusToSolidStatus = STATUS_WARNING
	Case VS_LOCAL_CRL_NOT_FOUND SignStatusToSolidStatus = STATUS_WARNING
	Case VS_CRL_TIME_EXPIRIED SignStatusToSolidStatus = STATUS_WARNING
	Case VS_REVOCATION_STATUS_UNKNOWN SignStatusToSolidStatus = STATUS_WARNING
	Case VS_REVOCATION_OCSP_ERROR SignStatusToSolidStatus = STATUS_WARNING
	Case VS_CTL_TIME_EXPIRIED SignStatusToSolidStatus = STATUS_WARNING

	Case Else SignStatusToSolidStatus = STATUS_BAD

	End Select
End Function

Function VerifyCertificate( ByRef oCert, ByVal iPolicyType )
	oCert.Profile = g_oProfile
	VerifyCertificate = oCert.IsValid( iPolicyType )
End Function

Function VerifySignatures( ByRef oSignatures, ByVal iLevel )
	Dim iResStatus : iResStatus = STATUS_UNKNOWN

	Dim oSignature, iCurStatus, i, iCount : iCount = oSignatures.Count
	For i = 0 to iCount - 1
		Set oSignature = oSignatures.Item(i)
		iCurStatus = oSignature.Verify( VF_ALL_POSSIBLE )
		iCurStatus = SignStatusToSolidStatus( iCurStatus )

		if iResStatus < iCurStatus then
			iResStatus = iCurStatus
		end if

		if SIGNATURE_TYPE_BASIC = oSignature.SignatureType then
			iCurStatus = VerifyCertificate( oSignature.Certificate, POLICY_TYPE_SIGNATURE )
			iCurStatus = SignStatusToSolidStatus( iCurStatus )

			if iResStatus < iCurStatus then
				iResStatus = iCurStatus
			end if
		end if

		if oSignature.Cosignature.Count > 0 then
			iCurStatus = VerifySignatures( oSignature.Cosignature, iLevel + 1 )
			iCurStatus = SignStatusToSolidStatus( iCurStatus )

			if iResStatus < iCurStatus then
				iResStatus = iCurStatus
			end if
		end if
	Next

	VerifySignatures = iResStatus
End Function

Function VerifySignature( ByVal sInputFilename )
	Dim sStatus, iStatus : iStatus = STATUS_INVALID

	On Error Resume Next

	g_oPKCS7Message.Load DT_SIGNED_DATA, sInputFilename
	if 0 = Err.Number then
		iStatus = VerifySignatures( g_oPKCS7Message.Signatures, 0 )
	end if

	Select Case iStatus
	Case STATUS_OK sStatus = "Успех"
	Case STATUS_WARNING sStatus = "Нет полного доверия к одной или нескольким ЭЦП"
	Case STATUS_BAD sStatus = "Одна или несколько ЭЦП недействительны!"
	Case Else sStatus = "Не является ЭЦП (при проверке возникла ошибка)!"
	End Select

	LogMsg "Совокупный статус ЭЦП файла """ & sInputFilename & """: " & sStatus

	On Error Goto 0
End Function

' ---
Function EncryptFile( ByVal sInputFilename, ByVal sOutputFilename )
	g_oPKCS7Message.Load DT_PLAIN_DATA, sInputFilename
	g_oPKCS7Message.Encrypt
	g_oPKCS7Message.Save DT_ENVELOPED_DATA, g_oProfile.EncryptExitFormat, sOutputFilename

	LogMsg "Файл """ & sInputFilename & """ зашифрован и сохранен в """ & sOutputFilename & """"
End Function

' ---
Function DecryptFile( ByVal sInputFilename, ByVal sOutputFilename )
	On Error Resume Next

	g_oPKCS7Message.Load DT_ENVELOPED_DATA, sInputFilename

	if Err.Number <> 0 then
		LogMsg "Файл """ & sInputFilename & """ не является шифрованным файлом!"
	elseif IsNothing(g_oPKCS7Message.Decrypt) then
		LogMsg "Файл """ & sInputFilename & """ не удается расшифровать!"
	else
		g_oPKCS7Message.Save DT_PLAIN_DATA, DER_TYPE, sOutputFilename

		LogMsg "Файл """ & sInputFilename & """ расшифрован и сохранен в """ & sOutputFilename & """"

		CachePinCode
	end if

	On Error Goto 0
End Function

' ---
Function ProcessFile( ByVal sInputFilename, ByVal sOutputFilename )

	if SIGN_WIZARD_TYPE = g_iOperationMode then
		ProcessFile = SignFile( sInputFilename, sOutputFilename & "." & g_sSignExt )
	elseif DROP_SIGNATURE_WIZARD_TYPE = g_iOperationMode then
		ProcessFile = TakeOffSignature( sInputFilename, TruncateFileExtension(sOutputFilename) )
	elseif VERIFY_SIGNATURE_WIZARD_TYPE = g_iOperationMode then
		ProcessFile = VerifySignature( sInputFilename )
	elseif ENCRYPT_WIZARD_TYPE = g_iOperationMode then
		ProcessFile = EncryptFile( sInputFilename, sOutputFilename & "." & g_sEncrExt )
	elseif DECRYPT_WIZARD_TYPE = g_iOperationMode then
		ProcessFile = DecryptFile( sInputFilename, TruncateFileExtension(sOutputFilename) )
	else
		PrintError "Указана неподдерживаемая операция, id=" & CStr(g_iOperationMode), -1
	end if
End Function

' ---
Function ProcessFolder( ByVal sInputFolder, ByVal sOutputFolder )
	if not oFSOShared.FolderExists( sInputFolder ) then
		PrintError "Не найден каталог '" & CStr(sInputFolder) & "'", -1
	end if
	if not oFSOShared.FolderExists( sOutputFolder ) then
		oFSOShared.CreateFolder( sOutputFolder )
	end if

	Dim oInputFolder : Set oInputFolder = oFSOShared.GetFolder( sInputFolder )
	Dim oOutputFolder : Set oOutputFolder = oFSOShared.GetFolder( sOutputFolder )

	' processing files
	Dim oFile, oFiles : Set oFiles = oInputFolder.Files
	For Each oFile in oFiles
		ProcessFile sInputFolder & "\" & oFile.Name, sOutputFolder & "\" & oFile.Name
	Next

	' processing subfolders
	Dim oFolder, oFolders : Set oFolders = oInputFolder.SubFolders
	For Each oFolder in oFolders
		ProcessFolder sInputFolder & "\" & oFolder.Name, sOutputFolder & "\" & oFolder.Name
	Next
End Function

' ===== main part =====
if WScript.Arguments.Count < 4 then
	PrintUsage(1)
else
	InitLog

	g_iOperationMode = ParseCryptoOperation( WScript.Arguments(3) )

	Set g_oProfile = LoadProfileByIdOrName( WScript.Arguments(2) )
	'g_oProfile.Display
	if IsNothing(g_oProfile) then
		PrintError "Не найдена настройка КриптоАРМ """ & CStr(WScript.Arguments(2)) & """", -1
	end if

	VerifyProfile g_oProfile, g_iOperationMode

	g_oPKCS7Message.Profile = g_oProfile
	g_sSignExt = g_oProfile.SignatureExtension(g_oProfile.SignExitFormat) ' since TD 4.2
	g_sEncrExt = g_oProfile.EncryptedExtension(g_oProfile.EncryptExitFormat) ' since TD 4.2

	ProcessFolder WScript.Arguments(0), WScript.Arguments(1)

	PrintError "Выполнение скрипта завершено успешно", 0
end if
