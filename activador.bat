@echo off
title Activacion de Windows 10, 11, Server 2019, Office 2019
color f2
:Principal
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   ::                   Activacion                   ::
echo   ::                 Menu Principal                 ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - Menu de Windows 10                ::
echo   :: Presione 2 - Menu de Windows 11                ::
echo   :: Presione 3 - Menu de Windows Server 2019       ::
echo   :: Presione 4 - Menu de Office 2019 y 2021        ::
echo   :: Presione 5 - Cargar Servidor y Validar Windows ::
echo   :: Presione 6 - Para Salir                        ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 123456 >nul
SET PRINOPTION=%ERRORLEVEL%
IF %PRINOPTION% EQU 1 (
	goto :Process
)
IF %PRINOPTION% EQU 2 (
	goto :Process11
)
IF %PRINOPTION% EQU 3 (
	goto :MENUWINDOWSSERVER
)
IF %PRINOPTION% EQU 4 (
	goto :OFFICE
)
IF %PRINOPTION% EQU 5 (
	goto :ACTIVATIONWINDOWS
)
IF %PRINOPTION% EQU 6 (
Echo Saliendo del sistema
pause
exit
)
::Apartir de esta parte comienza el menu de Windows 10
:Process
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Activacion de Windows 10                       ::
echo   :: Menu de Windows 10                             ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - Eleja su version de Windows 10    ::
echo   :: Presione 2 - Volver al Menu Principal          ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 12 >nul
SET WINOPTION=%ERRORLEVEL%
IF %WINOPTION% EQU 1 (
	goto :VERMENU
)
IF %WINOPTION% EQU 2 (
Echo Volviendo al Menu Principal
pause
goto :Principal
)
:VERMENU
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Eleja su Edicion de Windows 10                 ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Home                              ::
	echo   :: Presione 2 - Education                         ::
	echo   :: Presione 3 - Pro                               ::
	echo   :: Presione 4 - Enterprise                        ::
	echo   :: Presione 5 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 12345 >nul
	SET VERSIONWIN=%ERRORLEVEL%
	IF %VERSIONWIN% EQU 1 (
		goto :MENUHOME
	)
	IF %VERSIONWIN% EQU 2 (
		goto :MENUEDUCATION
	)
	IF %VERSIONWIN% EQU 3 (
		goto :MENUPRO
	)
	IF %VERSIONWIN% EQU 4 (
		goto :MENUENTERPRISE
	)
	IF %VERSIONWIN% EQU 5 (
		goto :Process
	)
:MENUHOME
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Home                  ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Home                              ::
	echo   :: Presione 2 - Home N                            ::
	echo   :: Presione 3 - Home Single Language              ::
	echo   :: Presione 4 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234 >nul
	SET HOME=%ERRORLEVEL%
	IF %HOME% EQU 1 (
	Echo Esta edicion no necesita clave, por favor proceda con la opcion 2 del menu principal.
	pause
	goto :Process
	)
	IF %HOME% EQU 2 (
	Echo Esta edicion no necesita clave, por favor proceda con la opcion 2 del menu principal.
	pause
	goto :Process
	)
	IF %HOME% EQU 3 (
	Echo Esta edicion no necesita clave, por favor proceda con la opcion 2 del menu principal.
	pause
	goto :Process
	)
	IF %HOME% EQU 4 (
	goto :VERMENU
		)
:MENUEDUCATION
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Education             ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Education                         ::
	echo   :: Presione 2 - Education N                       ::
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123 >nul
	SET EDU=%ERRORLEVEL%
	IF %EDU% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Education.
	slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
	goto :ACTIVATIONWINDOWS
	pause
	goto :Process
	)
	IF %EDU% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Education N.
	slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
	goto :ACTIVATIONWINDOWS
	pause
	goto :Process
	)
	IF %EDU% EQU 3 (
	goto :VERMENU
	)
:MENUPRO
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Pro                   ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Pro                               ::
	echo   :: Presione 2 - Pro N                             ::
	echo   :: Presione 3 - Pro Education                     ::
	echo   :: Presione 4 - Pro Education N                   ::
	echo   :: Presione 5 - Pro para Workstations             ::
	echo   :: Presione 6 - Pro para Workstations N           ::
	echo   :: Presione 7 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234567 >nul
	SET PRO=%ERRORLEVEL%
	IF %PRO% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Pro .
	slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Pro N.
	slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 3 (
	Echo Procederemos a introducir la clave de edicion Pro Education.
	slmgr /ipk 6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 4 (
	Echo Procederemos a introducir la clave de edicion Pro Education N.
	slmgr /ipk YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 5 (
	Echo Procederemos a introducir la clave de edicion Pro para Workstations.
	slmgr /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 6 (
	Echo Procederemos a introducir la clave de edicion Pro para Workstations N.
	slmgr /ipk 9FNHH-K3HBT-3W4TD-6383H-6XYWF
	goto :ACTIVATIONWINDOWS
	)	
	IF %PRO% EQU 7 (
	goto :VERMENU
	)
:MENUENTERPRISE
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Enterprise            ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Enterprise                        ::
	echo   :: Presione 2 - Enterprise N                      ::
	echo   :: Presione 3 - Enterprise G                      ::
	echo   :: Presione 4 - Enterprise G N                    ::
	echo   :: Presione 5 - Enterprise LTSC                   ::
	echo   :: Presione 6 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123456 >nul
	SET ENTERP=%ERRORLEVEL%
	IF %ENTERP% EQU 1 (
	Echo Procederemos a introducir la clave de edicion para Enterprise.
	slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 2 (
	Echo Procederemos a introducir la clave de edicion para Enterprise N.
	slmgr /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 3 (
	Echo Procederemos a introducir la clave de edicion para Enterprise G.
	slmgr /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 4 (
	Echo Procederemos a introducir la clave de edicion para Enterprise G N.
	slmgr /ipk 44RPN-FTY23-9VTTB-MP9BX-T84FV
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 5 (
	Echo Procederemos a introducir la clave de edicion para Enterprise LTSC.
	slmgr /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 6 (
	goto :VERMENU
	)
::Apartir de esta parte comienza el menu de Windows 11
:Process11
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Activacion de Windows 11                       ::
echo   :: Menu de Windows 11                             ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - Eleja su version de Windows 11    ::
echo   :: Presione 2 - Volver al Menu Principal          ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 12 >nul
SET WINOPTION=%ERRORLEVEL%
IF %WINOPTION% EQU 1 (
	goto :VERMENU11
)
IF %WINOPTION% EQU 2 (
Echo Volviendo al Menu Principal
pause
goto :Principal
)
:VERMENU11
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Eleja su Edicion de Windows 11                 ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Home                              ::
	echo   :: Presione 2 - Education                         ::
	echo   :: Presione 3 - Pro                               ::
	echo   :: Presione 4 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234 >nul
	SET VERSIONWIN=%ERRORLEVEL%
	IF %VERSIONWIN% EQU 1 (
		goto :MENUHOME11
	)
	IF %VERSIONWIN% EQU 2 (
		goto :MENUEDUCATION11
	)
	IF %VERSIONWIN% EQU 3 (
		goto :MENUPRO11
	)
	IF %VERSIONWIN% EQU 4 (
		goto :Process11
	)
:MENUHOME11
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Home                  ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Home                              ::
	echo   :: Presione 2 - Home Single Language              ::
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123 >nul
	SET HOME=%ERRORLEVEL%
	IF %HOME% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Home.
	slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
	goto :ACTIVATIONWINDOWS
	pause
	goto :Process11
	)
	IF %HOME% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Home Single Language.
	slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
	goto :ACTIVATIONWINDOWS
	pause
	goto :Process11
	)
	IF %HOME% EQU 3 (
	goto :VERMENU11
		)
:MENUEDUCATION11
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Education             ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Education                         ::
	echo   :: Presione 2 - Education N                       ::
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123 >nul
	SET EDU=%ERRORLEVEL%
	IF %EDU% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Education.
	slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
	goto :ACTIVATIONWINDOWS
	pause
	goto :Process11
	)
	IF %EDU% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Education N.
	slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
	goto :ACTIVATIONWINDOWS
	pause
	goto :Process11
	)
	IF %EDU% EQU 3 (
	goto :VERMENU11
	)
:MENUPRO11
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Pro                   ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Pro                               ::
	echo   :: Presione 2 - Pro N                             ::
	echo   :: Presione 3 - Pro para Workstations             ::
	echo   :: Presione 4 - Pro para Workstations N           ::
	echo   :: Presione 5 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 12345 >nul
	SET PRO=%ERRORLEVEL%
	IF %PRO% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Pro .
	slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Pro N.
	slmgr /ipk MH37W-N47XK-V7XM9-C7227-GCQG9
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 3 (
	Echo Procederemos a introducir la clave de edicion Pro para Workstations.
	slmgr /ipk NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 4 (
	Echo Procederemos a introducir la clave de edicion Pro para Workstations N.
	slmgr /ipk 9FNHH-K3HBT-3W4TD-6383H-6XYWF
	goto :ACTIVATIONWINDOWS
	)	
	IF %PRO% EQU 5 (
	goto :VERMENU11
	)
:: Apartir de esta parte comienza el menu de Windows Server 2019
:MENUWINDOWSSERVER
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Activacion de Windows Server 2019              ::
echo   :: Menu de Windows Server 2019                    ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - Eleja su version de Windows       ::
echo   :: Presione 2 - Volver al Menu Principal          ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 12 >nul
SET WINSERVEROPTION=%ERRORLEVEL%
IF %WINSERVEROPTION% EQU 1 (
	goto :VERMENUWINDOWSSERVER
)
IF %WINSERVEROPTION% EQU 2 (
	Echo Volviendo al Menu Principal
	pause
	goto :Principal
)
:VERMENUWINDOWSSERVER
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Eleja su Edicion de Windows Server 2019        ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Datacenter                        ::
	echo   :: Presione 2 - Standard                          ::
	echo   :: Presione 3 - Essentials                        ::
	echo   :: Presione 4 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234 >nul
	SET VERSIONWINSERVER=%ERRORLEVEL%
	IF %VERSIONWINSERVER% EQU 1 (
		goto :MENUDATACENTER
	)
	IF %VERSIONWINSERVER% EQU 2 (
		goto :MENUSTANDARD
	)
	IF %VERSIONWINSERVER% EQU 3 (
		goto :MENUESSENTIALS
	)
	IF %VERSIONWINSERVER% EQU 4 (
		goto :MENUWINDOWSSERVER
	)
:MENUDATACENTER
	cls>nul
	cls
	Echo Procederemos a introducir la clave de edicion Datacenter.
	slmgr /ipk WMDGN-G9PQG-XVVXX-R3X43-63DFG
	goto :ACTIVATION
:MENUSTANDARD
	cls>nul
	cls
	Echo Procederemos a introducir la clave de edicion Standard.
	slmgr /ipk N69G4-B89J2-4G8F4-WWYCC-J464C
	goto :ACTIVATION
:MENUESSENTIALS
	cls>nul
	cls
	Echo Procederemos a introducir la clave de edicion Essentials.
	slmgr /ipk WVDHN-86M7X-466P6-VHXV7-YY726
	goto :ACTIVATION
::Apartir de esta parte comienza el menu de Office 2019
:OFFICE
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Activacion de Office 2019 y 2021               ::
echo   :: Sub Menu                                       ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - Elija su paquete de Office 2019   ::
echo   :: Presione 2 - Elija su paquete de Office 2021   ::
echo   :: Presione 3 - Volver al Menu Principal          ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 123 >nul
SET OFFICEOPTION=%ERRORLEVEL%
IF %OFFICEOPTION% EQU 1 (
	goto :VERMENUOFFICE19
)
IF %OFFICEOPTION% EQU 2 (
	goto :VERMENUOFFICE21
)
IF %OFFICEOPTION% EQU 3 (
	Echo Volviendo al Menu Principal
	pause
	goto :Principal
)
:VERMENUOFFICE19
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Elija su paquete de Office 2019                ::
echo   :: Sub Menu                                       ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - ProPlus 2019                      ::
echo   :: Presione 2 - Project 2019                      ::
echo   :: Presione 3 - Visio 2019                        ::
echo   :: Presione 4 - Voler al Menu anterior            ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 1234 >nul
SET EDITOFFICEOPTION=%ERRORLEVEL%
IF %EDITOFFICEOPTION% EQU 1 (
	Echo Procederemos a introducir la clave de ProPlus 2019.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul&set i=1
	goto :ACTIVATIONOFFICE19
)
IF %EDITOFFICEOPTION% EQU 2 (
	Echo Procederemos a introducir la clave de Project 2019.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:PKD2B >nul&cscript //nologo ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B >nul&set i=1
	goto :ACTIVATIONOFFICE19
)
IF %EDITOFFICEOPTION% EQU 3 (
	Echo Procederemos a introducir la clave de Visio 2019.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\VisioPro2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\VisioPro2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:7VCBB >nul&cscript //nologo ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB >nul&set i=1
	goto :ACTIVATIONOFFICE19
)
IF %EDITOFFICEOPTION% EQU 4 (
	Echo Volviendo al Menu de Office 2019 y 2021
	pause
	goto :OFFICE
)
:VERMENUOFFICE21
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Elija su paquete de Office 2021                ::
echo   :: Sub Menu                                       ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - ProPlus 2021                      ::
echo   :: Presione 2 - Project 2021                      ::
echo   :: Presione 3 - Visio 2021                        ::
echo   :: Presione 4 - Voler al Menu anterior            ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 1234 >nul
SET EDITOFFICEOPTION=%ERRORLEVEL%
IF %EDITOFFICEOPTION% EQU 1 (
	Echo Procederemos a introducir la clave de ProPlus 2019.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6F7TH >nul&cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul&set i=1
	goto :ACTIVATIONOFFICE21
)
IF %EDITOFFICEOPTION% EQU 2 (
	Echo Procederemos a introducir la clave de Project 2019.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:QV9H8 >nul&cscript //nologo ospp.vbs /inpkey:KFTNWT-C6WBT-8HMGF-K9PRX-QV9H8 >nul&set i=1
	goto :ACTIVATIONOFFICE21
)
IF %EDITOFFICEOPTION% EQU 3 (
	Echo Procederemos a introducir la clave de Visio 2019.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\VisioPro2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\VisioPro2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:K2HT4 >nul&cscript //nologo ospp.vbs /inpkey:KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4 >nul&set i=1
	goto :ACTIVATIONOFFICE21
)
IF %EDITOFFICEOPTION% EQU 4 (
	Echo Volviendo al Menu de Office 2019 y 2021
	pause
	goto :OFFICE
)
::Apartir de esta parte comienza el menu para cargar el Servidor KMS y realizar la Activacion

:ACTIVATIONOFFICE19
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo ============================================================================&choice /n /c YN /m "Desea realizar otra operación [Y,N]?" & if errorlevel 2 exit) || (echo La conexion al servidor KMS fallo! Intentando conectarme a otro ... & echo Por favor espera... & echo. & echo. & set /a i+=1 & goto ACTIVATIONOFFICE)
pause
:ACTIVATIONOFFICE21
if %i%==1 set KMS_Sev=sethst:s8.uk.to
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo ============================================================================&choice /n /c YN /m "Desea realizar otra operación [Y,N]?" & if errorlevel 2 exit) || (echo La conexion al servidor KMS fallo! Intentando conectarme a otro ... & echo Por favor espera... & echo. & echo. & set /a i+=1 & goto ACTIVATIONOFFICE)
pause
:ACTIVATIONWINDOWS
cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar el servidor deseado                ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Servidor 1                        ::
	echo   :: Presione 2 - Servidor 2                        ::
	echo   :: Presione 3 - Servidor 3                        ::
	echo   :: Presione 4 - Servidor 4                        ::
	echo   :: Presione 5 - Volver al Menu Principal          ::		
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 12345 >nul
	SET SERVID=%ERRORLEVEL%
	IF %SERVID% EQU 1 (
	Echo Procederemos a introducir el servidor kms7.msguides.com.
	::cscript //nologo slmgr /skms kms7.msguides.com> nul&echo ============================================================================&echo.&echo.
	slmgr /skms kms7.msguides.com
	echo ============================================================================&echo.&echo.
	slmgr /ato
	goto :ACTIVATIONWINDOWS
		)
	IF %SERVID% EQU 2 (
	Echo Procederemos a introducir el servidor kms8.msguides.com.
	slmgr /skms kms8.msguides.com
	echo ============================================================================&echo.&echo.
	slmgr /ato
	goto :ACTIVATIONWINDOWS
	)
	IF %SERVID% EQU 3 (
	Echo Procederemos a introducir el servidor kms9.msguides.com.
	slmgr /skms kms9.msguides.com
	echo ============================================================================&echo.&echo.
	slmgr /ato
	goto :ACTIVATIONWINDOWS
	)
	IF %SERVID% EQU 4 (
	Echo Procederemos a introducir el servidor kms.digiboy.ir 
	slmgr /skms kms.digiboy.ir 
	echo ============================================================================&echo.&echo.
	slmgr /ato
	goto :ACTIVATIONWINDOWS
	)
	IF %SERVID% EQU 5 (
	goto :Principal
	)
:notsupported
echo.&echo ============================================================================&echo Lo siento! Tu versión no es compatible.
endlocal
