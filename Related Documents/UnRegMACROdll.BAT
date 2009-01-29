
REM Time Zone component
regsvr32 "..\Business Components\MACROTimeZoneBS30.dll" /u /s

REM Lock Components
regsvr32 "..\Data Components\MACROLOCKDS30.dll"  /u /s
regsvr32 "..\Business Components\MACROLOCKBS30.dll" /u /s 

REM User Components
regsvr32 "..\Business Components\MACROUserBS30.dll"  /u /s

REM Data Browser Components
regsvr32 "..\Business Components\MACRODBBS30.dll" /u /s

REM Data Entry Components
regsvr32 "..\State Components\MACRODEST30.dll" /u /s
regsvr32 "..\Data Components\MACRODEDS30.dll" /u /s 
regsvr32 "..\Business Components\MACRODEBS30.dll" /u /s 

REM MIMessage Components
regsvr32 "..\State Components\MACROMIMsgSt30.dll" /u /s
regsvr32 "..\Data Components\MACROMIMsgDS30.dll" /u /s 
regsvr32 "..\Business Components\MACROMIMsgBS30.dll" /u /s

rem System DataTransfer Components
regsvr32 "..\Business Components\MACROSysDataXfer30.dll" /u /s

REM AREZZOtoJScript Components
regsvr32 "..\Business Components\MACROAZJSBS30.DLL" /u /s

REM Subject Cache
regsvr32 "..\Subject Cache\MACROSCM30.DLL" /u /s
regsvr32 "..\Subject Cache\MACROSCI30.DLL" /u /s

REM Lock/Freeze DLL
regsvr32 "..\Business Components\MACROLFBS30.dll" /u /s

REM Web
regsvr32 "..\IO Components\MACROWWWIO30.DLL" /u /s
regsvr32 "..\IO Components\MACROIOEform30.DLL" /u /s

REM User Controls
regsvr32 "..\User Controls\MACROOptionGroup30.ocx" /u /s

REM Oc Discrepancy Component
regsvr32 "..\MACROOCDISC30.DLL" /u /s

REM Registration Component
regsvr32 "..\MACRORR30.DLL" /u /s

REM Version checking component
regsvr32 "..\Business Components\MACROVersion.DLL" /u /s

REM Active Directory
regasm /tlb:"IMEDActiveDirectory.tlb" "IMEDActiveDirectory.dll" /u /s

REM CCBS30
regasm /tlb:"..\Clinical Coding\MACROCCBS30\bin\Release\MACROCCBS30.tlb" "..\Clinical Coding\MACROCCBS30\bin\Release\MACROCCBS30.dll" /u /s

REM PluginInterface
regasm /tlb:"..\Clinical Coding\PluginInterface\bin\Release\PluginInterface.tlb" "..\Clinical Coding\PluginInterface\bin\Release\PluginInterface.dll" /u /s

pause
