'// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// NAME:        LSI.vbs
'//
'// Original code from:    http://www.cluberti.com/blog
'// Last Update: 28th December 2009
'//
'// Updated Code from: https://github.com/giefko
'// Last Update: 9th June 2017
'// Comment:     VBS example file for use as an OS info gathering template.
'//
'// NOTE:        Provided as-is - usage of this source assumes that you are at the
'//              very least familiar with the vbscript language being used and
'//              the tools used to create and debug this file.
'//
'//              In other words, if you break it, you get to keep the pieces.
'//
'//              Also, if you want to use this on W2K, prepare to hack, as this
'//              was really designed with XP+ systems in mind.
'//
'// ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// This script will require elevated privileges if run from a non-admin account, so
'// calling the ElevateThisScript() Sub should get the script a full admin token. This is
'// currently disabled, but if you need non-admin users to run this script, enable the
'// call to this subroutine to pop-up a dialog box (they'll of course need administrative
'// credentials to put in the challenge dialog before the script will execute with an
'// administrative token):

' ElevateThisScript()


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables and wMI connections for scripting against:
CONST HKEY_LOCAL_MACHINE = &H80000002
CONST SEARCH_KEY = "DigitalProductID"
Dim arrSubKeys(4,1)
Dim foundKeys
Dim iValues, arrDPID
foundKeys = Array()
iValues = Array()
arrSubKeys(0,0) = "Windows PID Key:       "
arrSubKeys(0,1) = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
arrSubKeys(2,0) = "Office XP PID Key:     "
arrSubKeys(2,1) = "SOFTWARE\Microsoft\Office\10.0\Registration"
arrSubKeys(1,0) = "Office 2003 PID Key:   "
arrSubKeys(1,1) = "SOFTWARE\Microsoft\Office\11.0\Registration"
arrSubKeys(3,0) = "Office 2007 PID Key:   "
arrSubKeys(3,1) = "SOFTWARE\Microsoft\Office\12.0\Registration"
arrSubKeys(4,0) = "Office 2010 PID Key:   "
arrSubKeys(4,1) = "SOFTWARE\Microsoft\Office\14.0\Registration\{10140000-0011-0000-1000-0000000FF1CE}"

strComputer = "."
Arch = ""
Sku = ""


Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
Set objSWbemDateTime = CreateObject("WbemScripting.SWbemDateTime")

Set colOSItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_OperatingSystem",,48)

Set colProcItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_Processor",,48)

Set colCompSysItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_ComputerSystem",,48) 

Set colTZItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_TimeZone",,48) 

Set colCompSysProdItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_ComputerSystemProduct",,48)

Set colBIOSItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_BIOS",,48)

Set colDiskItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_LogicalDisk",,48) 

Set colNetAdapConfigItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_NetworkAdapterConfiguration",,48) 

Set colVideoItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_VideoController",,48) 

Set colSoundItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_SoundDevice",,48) 


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Get OS SKU from Win32_OperatingSystem class:

    For Each objOSItem in colOSItems
    If objOSItem.BuildNumber => 6000 Then
        Arch = objOSItem.OSArchitecture

        Select Case objOSItem.OperatingSystemSKU
            Case 0 Sku = "Unknown Windows version"
            Case 1 Sku = "Ultimate Edition"
            Case 2 Sku = "Home Basic Edition"
            Case 3 Sku = "Home Premium Edition"
            Case 4 Sku = "Enterprise Edition"
            Case 5 Sku = "Home Basic N Edition"
            Case 6 Sku = "Business Edition"
            Case 7 Sku = "Standard Server Edition"
            Case 8 Sku = "Datacenter Server Edition"
            Case 9 Sku = "Small Business Server Edition"
            Case 10 Sku = "Enterprise Server Edition"
            Case 11 Sku = "Starter Edition"
            Case 12 Sku = "Datacenter Server Core Edition"
            Case 13 Sku = "Standard Server Core Edition"
            Case 14 Sku = "Enterprise Server Core Edition"
            Case 15 Sku = "Enterprise Server Edition for Itanium-Based Systems"
            Case 16 Sku = "Business N Edition"
            Case 17 Sku = "Web Server Edition"
            Case 18 Sku = "Cluster Server Edition"
            Case 19 Sku = "Home Server Edition"
            Case 20 Sku = "Storage Express Server Edition"
            Case 21 Sku = "Storage Standard Server Edition"
            Case 22 Sku = "Storage Workgroup Server Edition"
            Case 23 Sku = "Storage Enterprise Server Edition"
            Case 24 Sku = "Server For Small Business Edition"
            Case 25 Sku = "Small Business Server Premium Edition"
            Case Else
            Sku = "Could Not Determine Operating System SKU"
        End Select
    End If


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Get current OS locale setting from Win32_OperatingSystem class:

        Select Case objOSItem.Locale
            Case 0436 Locale = "Afrikaans (South Africa)"
            Case 041c Locale = "Albanian (Albania)"
            ' Case 045e Locale = "Amharic (Ethiopia)"
            Case 0401 Locale = "Arabic (Saudi Arabia)"
            Case 1401 Locale = "Arabic (Algeria)"
            Case 3c01 Locale = "Arabic (Bahrain)"
            Case 0c01 Locale = "Arabic (Egypt)"
            Case 0801 Locale = "Arabic (Iraq)"
            Case 2c01 Locale = "Arabic (Jordan)"
            Case 3401 Locale = "Arabic (Kuwait)"
            Case 3001 Locale = "Arabic (Lebanon)"
            Case 1001 Locale = "Arabic (Libya)"
            Case 1801 Locale = "Arabic (Morocco)"
            Case 2001 Locale = "Arabic (Oman)"
            Case 4001 Locale = "Arabic (Qatar)"
            Case 2801 Locale = "Arabic (Syria)"
            Case 1c01 Locale = "Arabic (Tunisia)"
            Case 3801 Locale = "Arabic (U.A.E.)"
            Case 2401 Locale = "Arabic (Yemen)"
            Case 042b Locale = "Armenian (Armenia)"
            Case 044d Locale = "Assamese"
            Case 082c Locale = "Azeri (Cyrillic)"
            Case 042c Locale = "Azeri (Latin)"
            Case 042d Locale = "Basque"
            Case 0423 Locale = "Belarusian"
            Case 0445 Locale = "Bengali (India)"
            Case 0845 Locale = "Bengali (Bangladesh)"
            Case 141A Locale = "Bosnian (Bosnia/Herzegovina)"
            Case 0402 Locale = "Bulgarian"
            Case 0455 Locale = "Burmese"
            Case 0403 Locale = "Catalan"
            Case 045c Locale = "Cherokee (United States)"
            Case 0804 Locale = "Chinese (PRC)"
            Case 1004 Locale = "Chinese (Singapore)"
            Case 0404 Locale = "Chinese (Taiwan)"
            Case 0c04 Locale = "Chinese (Hong Kong SAR)"
            Case 1404 Locale = "Chinese (Macao SAR)"
            Case 041a Locale = "Croatian"
            Case 101a Locale = "Croatian (Bosnia/Herzegovina)"
            Case 0405 Locale = "Czech"
            Case 0406 Locale = "Danish"
            Case 0465 Locale = "Divehi"
            Case 0413 Locale = "Dutch (Netherlands)"
            Case 0813 Locale = "Dutch (Belgium)"
            Case 0466 Locale = "Edo"
            Case 0409 Locale = "English (United States)"
            Case 0809 Locale = "English (United Kingdom)"
            Case 0c09 Locale = "English (Australia)"
            Case 2809 Locale = "English (Belize)"
            Case 1009 Locale = "English (Canada)"
            Case 2409 Locale = "English (Caribbean)"
            Case 3c09 Locale = "English (Hong Kong SAR)"
            Case 4009 Locale = "English (India)"
            Case 3809 Locale = "English (Indonesia)"
            Case 1809 Locale = "English (Ireland)"
            Case 2009 Locale = "English (Jamaica)"
            Case 4409 Locale = "English (Malaysia)"
            Case 1409 Locale = "English (New Zealand)"
            Case 3409 Locale = "English (Philippines)"
            Case 4809 Locale = "English (Singapore)"
            Case 1c09 Locale = "English (South Africa)"
            Case 2c09 Locale = "English (Trinidad)"
            Case 3009 Locale = "English (Zimbabwe)"
            Case 0425 Locale = "Estonian"
            Case 0438 Locale = "Faroese"
            Case 0429 Locale = "Farsi"
            Case 0464 Locale = "Filipino"
            Case 040b Locale = "Finnish"
            Case 040c Locale = "French (France)"
            Case 080c Locale = "French (Belgium)"
            Case 2c0c Locale = "French (Cameroon)"
            Case 0c0c Locale = "French (Canada)"
            Case 240c Locale = "French (DRC)"
            Case 300c Locale = "French (Cote d'Ivoire)"
            Case 3c0c Locale = "French (Haiti)"
            Case 140c Locale = "French (Luxembourg)"
            Case 340c Locale = "French (Mali)"
            Case 180c Locale = "French (Monaco)"
            Case 380c Locale = "French (Morocco)"
            Case e40c Locale = "French (North Africa)"
            Case 200c Locale = "French (Reunion)"
            Case 280c Locale = "French (Senegal)"
            Case 100c Locale = "French (Switzerland)"
            Case 1c0c Locale = "French (West Indies)"
            Case 0462 Locale = "Frisian (Netherlands)"
            Case 0467 Locale = "Fulfulde (Nigeria)"
            Case 042f Locale = "FYRO Macedonian"
            Case 083c Locale = "Gaelic (Ireland)"
            Case 043c Locale = "Gaelic (Scotland)"
            Case 0456 Locale = "Galician"
            Case 0437 Locale = "Georgian"
            Case 0407 Locale = "German (Germany)"
            Case 0c07 Locale = "German (Austria)"
            Case 1407 Locale = "German (Liechtenstein)"
            Case 1007 Locale = "German (Luxembourg)"
            Case 0807 Locale = "German (Switzerland)"
            Case 0408 Locale = "Greek"
            Case 0474 Locale = "Guarani (Paraguay)"
            Case 0447 Locale = "Gujarati"
            Case 0468 Locale = "Hausa (Nigeria)"
            Case 0475 Locale = "Hawaiian (United States)"
            Case 040d Locale = "Hebrew"
            Case 0439 Locale = "Hindi"
            ' Case 040e Locale = "Hungarian"
            Case 0469 Locale = "Ibibio (Nigeria)"
            Case 040f Locale = "Icelandic"
            Case 0470 Locale = "Igbo (Nigeria)"
            Case 0421 Locale = "Indonesian"
            Case 045d Locale = "Inuktitut"
            Case 0410 Locale = "Italian (Italy)"
            Case 0810 Locale = "Italian (Switzerland)"
            Case 0411 Locale = "Japanese"
            Case 044b Locale = "Kannada"
            Case 0471 Locale = "Kanuri (Nigeria)"
            Case 0860 Locale = "Kashmiri"
            Case 0460 Locale = "Kashmiri (Arabic)"
            Case 043f Locale = "Kazakh"
            Case 0453 Locale = "Khmer"
            Case 0457 Locale = "Konkani"
            Case 0412 Locale = "Korean"
            Case 0440 Locale = "Kyrgyz (Cyrillic)"
            Case 0454 Locale = "Lao"
            Case 0476 Locale = "Latin"
            Case 0426 Locale = "Latvian"
            Case 0427 Locale = "Lithuanian"
            ' Case 043e Locale = "Malay (Malaysia)"
            ' Case 083e Locale = "Malay (Brunei Darussalam)"
            Case 044c Locale = "Malayalam"
            Case 043a Locale = "Maltese"
            Case 0458 Locale = "Manipuri"
            Case 0481 Locale = "Maori (New Zealand)"
            ' Case 044e Locale = "Marathi"
            Case 0450 Locale = "Mongolian (Cyrillic)"
            Case 0850 Locale = "Mongolian (Mongolian)"
            Case 0461 Locale = "Nepali"
            Case 0861 Locale = "Nepali (India)"
            Case 0414 Locale = "Norwegian (Bokmål)"
            Case 0814 Locale = "Norwegian (Nynorsk)"
            Case 0448 Locale = "Oriya"
            Case 0472 Locale = "Oromo"
            Case 0479 Locale = "Papiamentu"
            Case 0463 Locale = "Pashto"
            Case 0415 Locale = "Polish"
            Case 0416 Locale = "Portuguese (Brazil)"
            Case 0816 Locale = "Portuguese (Portugal)"
            Case 0446 Locale = "Punjabi"
            Case 0846 Locale = "Punjabi (Pakistan)"
            Case 046B Locale = "Quecha (Bolivia)"
            Case 086B Locale = "Quecha (Ecuador)"
            Case 0C6B Locale = "Quecha (Peru)"
            Case 0417 Locale = "Rhaeto-Romanic"
            Case 0418 Locale = "Romanian"
            Case 0818 Locale = "Romanian (Moldava)"
            Case 0419 Locale = "Russian"
            Case 0819 Locale = "Russian (Moldava)"
            Case 043b Locale = "Sami (Lappish)"
            Case 044f Locale = "Sanskrit"
            Case 046c Locale = "Sepedi"
            Case 0c1a Locale = "Serbian (Cyrillic)"
            Case 081a Locale = "Serbian (Latin)"
            Case 0459 Locale = "Sindhi (India)"
            Case 0859 Locale = "Sindhi (Pakistan)"
            Case 045b Locale = "Sinhalese (Sri Lanka)"
            Case 041b Locale = "Slovak"
            Case 0424 Locale = "Slovenian"
            Case 0477 Locale = "Somali"
            ' Case 042e Locale = "Sorbian"
            Case 0c0a Locale = "Spanish (Spain - Modern Sort)"
            Case 040a Locale = "Spanish (Spain - Traditional Sort)"
            Case 2c0a Locale = "Spanish (Argentina)"
            Case 400a Locale = "Spanish (Bolivia)"
            Case 340a Locale = "Spanish (Chile)"
            Case 240a Locale = "Spanish (Colombia)"
            Case 140a Locale = "Spanish (Costa Rica)"
            Case 1c0a Locale = "Spanish (Dominican Republic)"
            Case 300a Locale = "Spanish (Ecuador)"
            Case 440a Locale = "Spanish (El Salvador)"
            Case 100a Locale = "Spanish (Guatemala)"
            Case 480a Locale = "Spanish (Honduras)"
            Case 580a Locale = "Spanish (Latin America)"
            Case 080a Locale = "Spanish (Mexico)"
            Case 4c0a Locale = "Spanish (Nicaragua)"
            Case 180a Locale = "Spanish (Panama)"
            Case 3c0a Locale = "Spanish (Paraguay)"
            Case 280a Locale = "Spanish (Peru)"
            Case 500a Locale = "Spanish (Puerto Rico)"
            Case 540a Locale = "Spanish (United States)"
            Case 380a Locale = "Spanish (Uruguay)"
            Case 200a Locale = "Spanish (Venezuela)"
            Case 0430 Locale = "Sutu"
            Case 0441 Locale = "Swahili"
            Case 041d Locale = "Swedish"
            Case 081d Locale = "Swedish (Finland)"
            Case 045a Locale = "Syriac"
            Case 0428 Locale = "Tajik"
            Case 045f Locale = "Tamazight (Arabic)"
            Case 085f Locale = "Tamazight (Latin)"
            Case 0449 Locale = "Tamil"
            Case 0444 Locale = "Tatar"
            Case 044a Locale = "Telugu"
            ' Case 041e Locale = "Thai"
            Case 0851 Locale = "Tibetan (Bhutan)"
            Case 0451 Locale = "Tibetan (PRC)"
            Case 0873 Locale = "Tigrigna (Eritrea)"
            Case 0473 Locale = "Tigrigna (Ethiopia)"
            Case 0431 Locale = "Tsonga"
            Case 0432 Locale = "Tswana"
            Case 041f Locale = "Turkish"
            Case 0442 Locale = "Turkmen"
            Case 0480 Locale = "Uighur (China)"
            Case 0422 Locale = "Ukrainian"
            Case 0420 Locale = "Urdu"
            Case 0820 Locale = "Urdu (India)"
            Case 0843 Locale = "Uzbek (Cyrillic)"
            Case 0443 Locale = "Uzbek (Latin)"
            Case 0433 Locale = "Venda"
            Case 042a Locale = "Vietnamese"
            Case 0452 Locale = "Welsh"
            Case 0434 Locale = "Xhosa"
            Case 0478 Locale = "Yi"
            Case 043d Locale = "Yiddish"
            Case 046a Locale = "Yoruba"
            Case 0435 Locale = "Zulu"
            Case 04ff Locale = "HID (Human Interface Device)"
            Case Else
            Locale = "Could Not Determine OS Locale"
        End Select


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_OperatingSystem class:

        Caption = objOSItem.Caption
        CSDVersion = objOSItem.CSDVersion
        CSName = objOSItem.CSName
        Version = objOSItem.Version
        BuildType = objOSItem.BuildType
        BuildNumber = objOSItem.BuildNumber
        SerialNumber = objOSItem.SerialNumber

        objSWbemDateTime.Value = objOSItem.InstallDate
        InstallDate = objSWbemDateTime.GetVarDate(True)

        objSWbemDateTime.Value = objOSItem.LastBootUpTime
        LastBootUpTime = objSWbemDateTime.GetVarDate(True)

        Status = objOSItem.Status
    Next


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_ComputerSystem class:

    For Each objCompSysItem in colCompSysItems
        CurrentTimeZone = objCompSysItem.CurrentTimeZone
        DaylightInEffect = objCompSysItem.DaylightInEffect
        TotalMemory = FormatNumber(objCompSysItem.TotalPhysicalMemory/1024^3, 2)
    Next


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_TimeZone class:

    For Each objTZItem in colTZItems
        TZName = objTZItem.StandardName
    Next


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_ComputerSystemProduct class:

    For Each objCompSysProdItem in colCompSysProdItems
        CompSysName = objCompSysProdItem.Name
        IdentifyingNumber = objCompSysProdItem.IdentifyingNumber
        UUID = objCompSysProdItem.UUID
    Next


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_BIOS class:

    For Each objBIOSItem in colBIOSItems
        SMBIOSVersion = objBIOSItem.SMBIOSBIOSVersion
    Next
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
set fso = CreateObject("Scripting.FileSystemObject")


Path = fso.GetParentFolderName(WScript.ScriptFullName)
Name="\System Information.txt" 

If Path <> "" And Name <> "" Then
	set tsFile = fso.OpenTextFile(Path & Name,2,True)
End If



'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Start echoing output to the screen

    tsFile.WriteLine  "     ------------------------------------------"
    tsFile.WriteLine "                   System Details"
    tsFile.WriteLine "     ------------------------------------------"
    tsFile.WriteLine ""
    tsFile.WriteLine "     Computer Name:        " & CSName
    tsFile.WriteLine ""
    tsFile.WriteLine ""
    tsFile.WriteLine "     Operating System Information:"
    tsFile.WriteLine "     ============================="
    tsFile.WriteLine "     Operating System:     " & Caption & Arch
    tsFile.WriteLine "     Version:               " & Version & " " & Sku & " " & CSDVersion
    tsFile.WriteLine "     Build Type:            " & BuildType
    tsFile.WriteLine "     Locale:                " & Locale
    tsFile.WriteLine "     Serial Number:         " & SerialNumber
    tsFile.WriteLine ""
    tsFile.WriteLine "     Current Time Zone:     " & TZName
    tsFile.WriteLine "     Offset from UTC:       " & CurrentTimeZone/60 & " hours"
    tsFile.WriteLine "     DST In Effect:         " & DaylightInEffect
    tsFile.WriteLine ""
'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Get the product keys (function at the end of script):

GetKeys()

    tsFile.WriteLine ""
    tsFile.WriteLine "     Install Date:          " & InstallDate
    tsFile.WriteLine "     Last Boot Time:        " & LastBootUpTime
    tsFile.WriteLine "     Local Date/Time:       " & Now()
    tsFile.WriteLine ""
    tsFile.WriteLine "     System Status:         " & Status
    tsFile.WriteLine ""
    tsFile.WriteLine ""
    tsFile.WriteLine "     Hardware Information:"
    tsFile.WriteLine "     ====================="


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set and echo variables gathered from Win32_Processor class:

    For Each objProcItem in colProcItems
        Select Case objProcItem.Architecture
            Case 0 CPUArch = "x86"
            Case 1 CPUArch = "MIPS"
            Case 2 CPUArch = "Alpha"
            Case 3 CPUArch = "PowerPC"
            Case 6 CPUArch = "Itanium"
            Case 9 CPUArch = "x64"
            Case Else
            CPUArch = "Could Not Determine CPU Architecture"
        End Select
    tsFile.WriteLine "     CPU:                  " & objProcItem.Name & " (" & CPUArch & ")"
    Next
    tsFile.WriteLine ""
    tsFile.WriteLine "     Physical Memory:      " & TotalMemory & " GB"
    tsFile.WriteLine ""


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_VideoController class:

    For Each objVideoItem in colVideoItems
        tsFile.WriteLine "     Video Card:           " & objVideoItem.Name

        If Not objVideoItem.AdapterDACType = "" Then
            tsFile.WriteLine "     Adapter DAC:           " & objVideoItem.AdapterDACType
        End if

        tsFile.WriteLine     "     PNP Device ID:         " & objVideoItem.PNPDeviceID

        If Not objVideoItem.AdapterRAM = "" Then
            tsFile.WriteLine "     Video RAM:             " & objVideoItem.AdapterRAM/1024^2 & " MB"
        End If

        tsFile.WriteLine     "     Driver Version:        " & objVideoItem.DriverVersion

        objSWbemDateTime.Value = objVideoItem.DriverDate
        DriverDate = objSWbemDateTime.GetVarDate(False)
        tsFile.WriteLine "     Driver Date:           " & DriverDate

        tsFile.WriteLine ""
    Next


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_SoundDevice class:

    For Each objSoundItem in colSoundItems
        tsFile.WriteLine "     Sound Card:           " & objSoundItem.Name
        tsFile.WriteLine "     Manufacturer:          " & objSoundItem.Manufacturer
        tsFile.WriteLine "     PNP Device ID:         " & objSoundItem.PNPDeviceID
        tsFile.WriteLine ""
    Next


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_LogicalDisk class:

    For Each objDiskItem in colDiskItems
        If objDiskItem.DriveType = 3 Then
            tsFile.WriteLine "     Volume:               " & objDiskItem.Caption
            tsFile.WriteLine "     Compressed:            " & objDiskItem.Compressed
            tsFile.WriteLine "     File System:           " & objDiskItem.FileSystem
            tsFile.WriteLine "     Volume Size:           " & FormatNumber(objDiskItem.Size/1024^3, 2) & " GB"
            tsFile.WriteLine "     Free Space:            " & FormatNumber(objDiskItem.FreeSpace/1024^3, 2) & " GB"
            tsFile.WriteLine ""
        End If
    Next


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Set variables gathered from Win32_NetworkAdapterConfiguration class:

    For Each objNetAdapConfigItem in colNetAdapConfigItems
        If isNull(objNetAdapConfigItem.IPAddress) Then
            '// Skip adapter, not currently used
        Else
            tsFile.WriteLine "     Network Adapter:      " & objNetAdapConfigItem.Description
            tsFile.WriteLine "     MAC Address:           " & objNetAdapConfigItem.MACAddress
            tsFile.WriteLine "     DHCP Enabled:          " & objNetAdapConfigItem.DHCPEnabled
            tsFile.WriteLine "     IP Address:            " & Join(objNetAdapConfigItem.IPAddress, ",")
            tsFile.WriteLine "     Subnet Mask:           " & Join(objNetAdapConfigItem.IPSubnet, ",")
            tsFile.WriteLine "     Default Gateway:       " & Join(objNetAdapConfigItem.DefaultIPGateway, ",")
            If objNetAdapConfigItem.DHCPEnabled = True Then

                objSWbemDateTime.Value = objNetAdapConfigItem.DHCPLeaseObtained
                DHCPLeaseObtained = objSWbemDateTime.GetVarDate(True)
                tsFile.WriteLine "     Lease Obtained:        " & DHCPLeaseObtained

                objSWbemDateTime.Value = objNetAdapConfigItem.DHCPLeaseExpires
                DHCPLeaseExpires = objSWbemDateTime.GetVarDate(True)
                tsFile.WriteLine "     Lease Exipres:         " & DHCPLeaseExpires

                tsFile.WriteLine "     DHCP Servers:          " & objNetAdapConfigItem.DHCPServer
            End If
            tsFile.WriteLine "     DNS Server:            " & Join(objNetAdapConfigItem.DNSServerSearchOrder, ",")
            If Not objNetAdapConfigItem.WINSPrimaryServer = "" Then
                tsFile.WriteLine "     WINS Primary Server:   " & objNetAdapConfigItem.WINSPrimaryServer
                If Not objNetAdapConfigItem.WINSSecondaryServer = "" Then
                    tsFile.WriteLine "     WINS Secondary Server: " & objNetAdapConfigItem.WINSPrimaryServer
                End If
                tsFile.WriteLine "     Enable LMHosts Lookup: " & objNetAdapConfigItem.WINSEnableLMHostsLookup
            End If
            tsFile.WriteLine ""
        End If
    Next
    tsFile.WriteLine ""
    tsFile.WriteLine "     System Information:"
    tsFile.WriteLine "     ==================="
    tsFile.WriteLine "     Computer:             " & CompSysName
    tsFile.WriteLine "     Serial Number:         " & IdentifyingNumber
    tsFile.WriteLine "     BIOS Version:          " & SMBIOSVersion
    tsFile.WriteLine "     UUID:                  " & UUID
    tsFile.WriteLine ""


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// This pause call is disabled, but you may wish to enable it if running this script
'// with the ElevateThisScript() subroutine call enabled above:

' PressEnter()


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Functions: GetKeys() and decodeKey(iValues, strProduct)
'//
'// Credit where credit is due - I found and modified script posted by user "Parabellum"
'// found on http://www.visualbasicscript.com/m42793.aspx, hacked it up a bit, and used
'// it here to post keys:

Public Function GetKeys()

    For x = LBound(arrSubKeys, 1) To UBound(arrSubKeys, 1)
    objReg.GetBinaryValue HKEY_LOCAL_MACHINE, arrSubKeys(x,1), SEARCH_KEY, arrDPIDBytes

        If Not IsNull(arrDPIDBytes) Then
            Call decodeKey(arrDPIDBytes, arrSubKeys(x,0))
        Else
            objReg.EnumKey HKEY_LOCAL_MACHINE, arrSubKeys(x,1), arrGUIDKeys

            If Not IsNull(arrGUIDKeys) Then
                For Each GUIDKey In arrGUIDKeys
                    objReg.GetBinaryValue HKEY_LOCAL_MACHINE, arrSubKeys(x,1) & "\" & GUIDKey, SEARCH_KEY, arrDPIDBytes
        
                    If Not IsNull(arrDPIDBytes) Then
                        Call decodeKey(arrDPIDBytes, arrSubKeys(x,0))
                    End If
                Next
            End If
        End If
    Next

End Function


    Public Function decodeKey(iValues, strProduct)

        Dim arrDPID
        arrDPID = Array()

            For i = 52 to 66
                ReDim Preserve arrDPID( UBound(arrDPID) + 1 )
                arrDPID( UBound(arrDPID) ) = iValues(i)
            Next

        Dim arrChars
        arrChars = Array("B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9")

            For i = 24 To 0 Step -1
                k = 0
        
                For j = 14 To 0 Step -1
                    k = k * 256 Xor arrDPID(j)
                    arrDPID(j) = Int(k / 24)
                    k = k Mod 24
                Next
        
                strProductKey = arrChars(k) & strProductKey
        
                If i Mod 5 = 0 And i <> 0 Then
                    strProductKey = "-" & strProductKey
                End If
            Next
    
        ReDim Preserve foundKeys( UBound(foundKeys) + 1 )
        foundKeys( UBound(foundKeys) ) = strProductKey
        strKey = UBound(foundKeys)
        tsFile.WriteLine "     " & strProduct & "" & foundKeys(strKey)
    End Function


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Subroutine: PressEnter()
'//
'// Adds a pause with a "Press the ENTER key to continue." message when called.
'//
'// Usage:  Call this Subroutine to get a pause that will clear when the user presses the
'// ENTER key (and ONLY the ENTER key) on their keyboard:

Sub PressEnter()
    tsFile.WriteLine ""
    strMessage = "Press the ENTER key to continue. "
    Wscript.StdOut.Write strMessage

    Do While Not WScript.StdIn.AtEndOfLine
        Input = WScript.StdIn.Read(1)
    Loop
End Sub


'//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'//
'// Subroutine: ElevateThisScript()
'//
'// Forces the currently running script to prompt for elevation if it detects that the
'// current user credentials do not have administrative privileges.
'//
'// If run on Windows XP / Server 2003 this script will cause the RunAs dialog to appear
'// if the user does not have administrative rights, giving the opportunity to run as an
'// administrator.
'//
'// If run on Windows Vista / Server 2008 this script will cause a UAC dialog to appear
'// if the user does not have administrative rights and UAC is enabled, giving the
'// opportunity for the script to run properly for a LUA user.
'//
'// This Sub Attempts to call the script with its original arguments.  Arguments that
'// contain a space will be wrapped in double quotes when the script calls itself again.
'//
'// Usage:  Add a call to this sub (ElevateThisScript) to the beginning of your script to
'// ensure that the script gets an administrative token.

Sub ElevateThisScript()

	Const HKEY_CLASSES_ROOT  = &H80000000
	Const HKEY_CURRENT_USER  = &H80000001
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const HKEY_USERS         = &H80000003
	const KEY_QUERY_VALUE	  = 1
	Const KEY_SET_VALUE		  = 2

	Dim scriptEngine, engineFolder, argString, arg, Args, scriptCommand
	Dim objShellApp : Set objShellApp = CreateObject("Shell.Application")

	scriptEngine = Ucase(Mid(Wscript.FullName,InstrRev(Wscript.FullName,"\")+1))
	engineFolder = Left(Wscript.FullName,InstrRev(Wscript.FullName,"\"))
	argString = ""

	Set Args = Wscript.Arguments

	For each arg in Args						'loop though argument array as a collection to rebuild argument string
		If instr(arg," ") > 0 Then arg = """" & arg & """"	'if the argument contains a space wrap it in double quotes
		argString = argString & " " & Arg
	Next

	scriptCommand = engineFolder & scriptEngine

	Dim strComputer : strComputer = "."

	Dim objReg, bHasAccessRight
	Set objReg=GetObject("winmgmts:"_
		& "{impersonationLevel=impersonate}!\\" &_
		strComputer & "\root\default:StdRegProv")

	'Check for administrative registry access rights
	objReg.CheckAccess HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\CrashControl", _
		KEY_SET_VALUE, bHasAccessRight

	If bHasAccessRight = True Then

		HasRequiredRegAccess = True
		Exit Sub

	Else

		HasRequiredRegAccess = False
		objShellApp.ShellExecute scriptCommand, " """ & Wscript.ScriptFullName & """" & argString, "", "runas"
		WScript.Quit
	End If
End Sub
tsFile.Close