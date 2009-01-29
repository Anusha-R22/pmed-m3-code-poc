
REM Time Zone component
regsvr32 "..\Business Components\MACROTimeZoneBS30.dll" 

REM Lock Components
regsvr32 "..\Data Components\MACROLOCKDS30.dll" 
regsvr32 "..\Business Components\MACROLOCKBS30.dll" 

REM User Components
regsvr32 "..\Business Components\MACROUserBS30.dll" 

REM Data Browser Components
regsvr32 "..\Business Components\MACRODBBS30.dll" 

REM Data Entry Components
regsvr32 "..\State Components\MACRODEST30.dll"
regsvr32 "..\Data Components\MACRODEDS30.dll" 
regsvr32 "..\Business Components\MACRODEBS30.dll" 

REM MIMessage Components
regsvr32 "..\State Components\MACROMIMsgSt30.dll"
regsvr32 "..\Data Components\MACROMIMsgDS30.dll" 
regsvr32 "..\Business Components\MACROMIMsgBS30.dll"

rem System DataTransfer Components
regsvr32 "..\Business Components\MACROSysDataXfer30.dll"

REM AREZZOtoJScript Components
regsvr32 "..\Business Components\MACROAZJSBS30.DLL"

REM Subject Cache
regsvr32 "..\Subject Cache\MACROSCM30.DLL"
regsvr32 "..\Subject Cache\MACROSCI30.DLL"

REM Lock/Freeze DLL
regsvr32 "..\Business Components\MACROLFBS30.dll"

REM Web
regsvr32 "..\IO Components\MACROWWWIO30.DLL"
regsvr32 "..\IO Components\MACROIOEform30.DLL"

REM User Controls
regsvr32 "..\User Controls\MACROOptionGroup30.ocx"

REM Oc Discrepancy Component
regsvr32 "..\MACROOCDISC30.DLL"

REM Registration Component
regsvr32 "..\MACRORR30.DLL"

REM IMED DLLs that MACRO is dependent on
regsvr32 "IMEDSettings10.dll"
regsvr32 "IMEDCheckSum10.dll"
regsvr32 "IMEDQD20.dll"
regsvr32 "IMEDQDSDS20.dll"
regsvr32 "IMEDQDS20.dll"
regsvr32 "IMEDQS20.dll"
regsvr32 "..\MACROAccess22ToMSDE22.dll"

REM Version checking component
regsvr32 "..\Business Components\MACROVersion.DLL" /s

REM Active Directory
regasm /tlb:"IMEDActiveDirectory.tlb" "IMEDActiveDirectory.dll"

REM CCBS30
regasm /tlb:"..\Clinical Coding\MACROCCBS30\bin\Release\MACROCCBS30.tlb" "..\Clinical Coding\MACROCCBS30\bin\Release\MACROCCBS30.dll"

REM PluginInterface
regasm /tlb:"..\Clinical Coding\PluginInterface\bin\Release\PluginInterface.tlb" "..\Clinical Coding\PluginInterface\bin\Release\PluginInterface.dll"

pause
