@echo off
title Activacion de Windows y Office
color f2
:INICIO
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   ::                      MarteTeam                         ::
echo   ::             Activador de Windows y Office              ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Este activador contiene los siguientes sistemas,       ::
echo   :: Windows 7, 8.1, 10, 11                                 ::  
echo   :: Server 2008, 2008 R2, 2012, 2012 R2, 2016, 2019, 2022  ::
echo   :: Office 2013, 2016, 2019, 2021                          ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
pause
:Principal
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   ::                      Activacion                      ::
echo   ::                    Menu Principal                    ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - Menu de Windows                         ::
echo   :: Presione 2 - Menu de Office                          ::
echo   :: Presione 3 - Cargar Servidor y Validar Windows       ::
echo   :: Presione 4 - Para Salir                              ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 1234 >nul
SET PRINOPTION=%ERRORLEVEL%
IF %PRINOPTION% EQU 1 (
	goto :VERMENWIN
)
IF %PRINOPTION% EQU 2 (
	goto :OFFICE
)
IF %PRINOPTION% EQU 3 (
	goto :ACTIVATIONWINDOWS
)
IF %PRINOPTION% EQU 4 (
Echo Saliendo del sistema
pause
exit
)
::Apartir de esta parte comienza los menus de Windows
:VERMENWIN
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Menu de Windows                                ::
	echo   :: Eleja su Windows                               ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Menu de Windows 7                 ::
	echo   :: Presione 2 - Menu de Windows 8.1               ::
	echo   :: Presione 3 - Menu de Windows 10                ::
	echo   :: Presione 4 - Menu de Windows 11                ::
	echo   :: Presione 5 - Menu de Windows Server            ::
	echo   :: Presione 6 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 12345678 >nul
	SET WINVER=%ERRORLEVEL%
	IF %WINVER% EQU 1 (
		goto :VERMENU7
	)
	IF %WINVER% EQU 2 (
		goto :VERMENU81
	)
	IF %WINVER% EQU 3 (
		goto :VERMENU10
	)
	IF %WINVER% EQU 4 (
		goto :VERMENU11
	)
	IF %WINVER% EQU 5 (
		goto :VERMENWINSERVER
	)
	IF %WINVER% EQU 6 (
		goto :Principal
	)
::Apartir de esta parte comienza el menu de Windows 7
:VERMENU7
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows 7                        ::
	echo   :: Eleja su Edicion de Windows 7                  ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Starter                           ::
	echo   :: Presione 2 - Home                              ::
	echo   :: Presione 3 - Pro                               ::
	echo   :: Presione 4 - Ultimate                          ::
	echo   :: Presione 5 - Enterprise                        ::
	echo   :: Presione 6 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123456 >nul
	SET VERSIONWIN=%ERRORLEVEL%
	IF %VERSIONWIN% EQU 1 (
		goto :MENUSTARTER
	)
	IF %VERSIONWIN% EQU 2 (
		goto :MENUHOME
	)
	IF %VERSIONWIN% EQU 3 (
		goto :MENUPRO
	)
	IF %VERSIONWIN% EQU 4 (
		goto :MENUULTIMATE
	)
	IF %VERSIONWIN% EQU 5 (
		goto :MENUENTERPRISE
	)
	IF %VERSIONWIN% EQU 6 (
		goto :VERMENWIN
	)
:MENUSTARTER
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Starter               ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Starter                           ::
	echo   :: Presione 2 - Starter N                         ::
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123 >nul
	SET STARTER=%ERRORLEVEL%
	IF %STARTER% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Starter.
	slmgr /ipk 7Q28W-FT9PC-CMMYT-WHMY2-89M6G
	goto :ACTIVATIONWINDOWS
	)
	IF %STARTER% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Starter N.
	slmgr /ipk D4C3G-38HGY-HGQCV-QCWR8-97FFR
	goto :ACTIVATIONWINDOWS
	)
	IF %STARTER% EQU 3 (
	goto :VERMENU7
	)
:MENUHOME
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Home                  ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Home Basic                        ::
	echo   :: Presione 2 - Home Basique N                    ::
	echo   :: Presione 3 - Home Premium                      ::
	echo   :: Presione 4 - Home Premium N                    ::
	echo   :: Presione 5 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234 >nul
	SET HOME=%ERRORLEVEL%
	IF %HOME% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Home Basic.
	slmgr /ipk YGFVB-QTFXQ-3H233-PTWTJ-YRYRV
	goto :ACTIVATIONWINDOWS
	)
	IF %HOME% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Home Basic N.
	slmgr /ipk MD83G-H98CG-DXPYQ-Q8GCR-HM8X2
	goto :ACTIVATIONWINDOWS
	)
	IF %HOME% EQU 3 (
	Echo Procederemos a introducir la clave de edicion Home Premium.
	slmgr /ipk HPQ2-RMFJH-74XYM-BH4JX-XM76F
	goto :ACTIVATIONWINDOWS
	)
	IF %HOME% EQU 4 (
	Echo Procederemos a introducir la clave de edicion Home Premium N.
	slmgr /ipk D3PVQ-V7M4J-9Q9K3-GG4K3-F99JM
	goto :ACTIVATIONWINDOWS
	)
	IF %HOME% EQU 5 (
	goto :VERMENU7
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
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123 >nul
	SET PRO=%ERRORLEVEL%
	IF %PRO% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Pro .
	slmgr /ipk HYF8J-CVRMY-CM74G-RPHKF-PW487
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Pro N.
	slmgr /ipk BKFRB-RTCT3-9HW44-FX3X8-M48M6
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 3 (
	goto :VERMENU7
	)
:MENUULTIMATE
	cls>nul
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Especificar cual edicion Ultimate              ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Ultimate                          ::
	echo   :: Presione 2 - Ultimate N                        ::
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123 >nul
	SET ULTIMATE=%ERRORLEVEL%
	IF %ULTIMATE% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Ultimate .
	slmgr /ipk D4F6K-QK3RD-TMVMJ-BBMRX-3MBMV
	goto :ACTIVATIONWINDOWS
	)
	IF %ULTIMATE% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Ultimate N.
	slmgr /ipk HTJK6-DXX8T-TVCR6-KDG67-97J8Q
	goto :ACTIVATIONWINDOWS
	)
	IF %ULTIMATE% EQU 3 (
	goto :VERMENU7
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
	echo   :: Presione 3 - Enterprise E                      ::
	echo   :: Presione 4 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234 >nul
	SET ENTERP=%ERRORLEVEL%
	IF %ENTERP% EQU 1 (
	Echo Procederemos a introducir la clave de edicion para Enterprise.
	slmgr /ipk H7X92-3VPBB-Q799D-Y6JJ3-86WC6
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 2 (
	Echo Procederemos a introducir la clave de edicion para Enterprise N.
	slmgr /ipk BQ4TH-BWRRY-424Y9-7PQX2-B4WBD
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 3 (
	Echo Procederemos a introducir la clave de edicion para Enterprise N.
	slmgr /ipk C29WB-22CC8-VJ326-GHFJW-H9DH4
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 4 (
	goto :VERMENU7
	)
::Apartir de esta parte comienza el menu de Windows 8.1
:VERMENU81
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows 8.1                      ::
	echo   :: Eleja su Edicion de Windows 8.1                ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Pro                               ::
	echo   :: Presione 2 - Enterprise                        ::
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123 >nul
	SET VERSIONWIN=%ERRORLEVEL%
	IF %VERSIONWIN% EQU 1 (
		goto :MENUPRO
	)
	IF %VERSIONWIN% EQU 2 (
		goto :MENUENTERPRISE
	)
	IF %VERSIONWIN% EQU 3 (
		goto :VERMENWIN
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
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234567 >nul
	SET PRO=%ERRORLEVEL%
	IF %PRO% EQU 1 (
	Echo Procederemos a introducir la clave de edicion Pro .
	slmgr /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Pro N.
	slmgr /ipk HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
	goto :ACTIVATIONWINDOWS
	)
	IF %PRO% EQU 3 (
	goto :VERMENU81
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
	echo   :: Presione 3 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 123456 >nul
	SET ENTERP=%ERRORLEVEL%
	IF %ENTERP% EQU 1 (
	Echo Procederemos a introducir la clave de edicion para Enterprise.
	slmgr /ipk MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 2 (
	Echo Procederemos a introducir la clave de edicion para Enterprise N.
	slmgr /ipk TT4HM-HN7YT-62K67-RGRQJ-JFFXW
	goto :ACTIVATIONWINDOWS
	)
	IF %ENTERP% EQU 3 (
	goto :VERMENU81
	)
::Apartir de esta parte comienza el menu de Windows 10
:VERMENU10
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows 10                       ::
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
		goto :VERMENWIN
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
	goto :Principal
	)
	IF %HOME% EQU 2 (
	Echo Esta edicion no necesita clave, por favor proceda con la opcion 2 del menu principal.
	pause
	goto :Principal
	)
	IF %HOME% EQU 3 (
	Echo Esta edicion no necesita clave, por favor proceda con la opcion 2 del menu principal.
	pause
	goto :Principal
	)
	IF %HOME% EQU 4 (
	goto :VERMENU10
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
	)
	IF %EDU% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Education N.
	slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
	goto :ACTIVATIONWINDOWS
	)
	IF %EDU% EQU 3 (
	goto :VERMENU10
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
	goto :VERMENU10
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
	goto :VERMENU10
	)
::Apartir de esta parte comienza el menu de Windows 11
:VERMENU11
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows 11                       ::
	echo   :: Eleja su Edicion de Windows 11                 ::
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
		goto :MENUHOME11
	)
	IF %VERSIONWIN% EQU 2 (
		goto :MENUEDUCATION11
	)
	IF %VERSIONWIN% EQU 3 (
		goto :MENUPRO11
	)
	IF %VERSIONWIN% EQU 4 (
		goto :MENUENTERPRISE11
	)
	IF %VERSIONWIN% EQU 4 (
		goto :VERMENWIN
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
	)
	IF %HOME% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Home Single Language.
	slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
	goto :ACTIVATIONWINDOWS
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
	)
	IF %EDU% EQU 2 (
	Echo Procederemos a introducir la clave de edicion Education N.
	slmgr /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
	goto :ACTIVATIONWINDOWS
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
:MENUENTERPRISE11
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
	goto :VERMENU11
	)
::Apartir de esta parte comienza los menus de Windows Server
:VERMENWINSERVER
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Menu de Windows Server                         ::
	echo   :: Eleja su Windows Server                        ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Menu de Windows Server 2008       ::
	echo   :: Presione 2 - Menu de Windows Server 2008 R2    ::
	echo   :: Presione 3 - Menu de Windows Server 2012       ::
	echo   :: Presione 4 - Menu de Windows Server 2012 R2    ::
	echo   :: Presione 5 - Menu de Windows Server 2016       ::
	echo   :: Presione 6 - Menu de Windows Server 2019       ::
	echo   :: Presione 7 - Menu de Windows Server 2022       ::
	echo   :: Presione 8 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 12345678 >nul
	SET WINVER=%ERRORLEVEL%
	IF %WINVER% EQU 1 (
		goto :VERMENUWINDOWSSERVER2008
	)
	IF %WINVER% EQU 2 (
		goto :VERMENUWINDOWSSERVER2008R2
	)
	IF %WINVER% EQU 3 (
		goto :VERMENUWINDOWSSERVER2012
	)
	IF %WINVER% EQU 4 (
		goto :VERMENUWINDOWSSERVER2012R2
	)
	IF %WINVER% EQU 5 (
		goto :VERMENUWINDOWSSERVER2016
	)
	IF %WINVER% EQU 6 (
		goto :VERMENUWINDOWSSERVER2019
	)
	IF %WINVER% EQU 7 (
		goto :VERMENUWINDOWSSERVER2022
	)
	IF %WINVER% EQU 8 (
		goto :VERMENWIN
	)
:: Apartir de esta parte comienza el menu de Windows Server 2008
:VERMENUWINDOWSSERVER2008
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows Server 2008              ::
	echo   :: Eleja su Edicion de Windows Server 2008        ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Datacenter                        ::
	echo   :: Presione 2 - Standard                          ::
	echo   :: Presione 3 - Enterprise                        ::
	echo   :: Presione 4 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234 >nul
	SET VERSIONWINSERVER=%ERRORLEVEL%
	IF %VERSIONWINSERVER% EQU 1 (
		Echo Procederemos a introducir la clave de edicion Datacenter.
		slmgr /ipk 7M67G-PC374-GR742-YH8V4-TCBY3
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 2 (
		Echo Procederemos a introducir la clave de edicion Standard.
		slmgr /ipk TM24T-X9RMF-VWXK6-X8JC9-BFGM2
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 3 (
		Echo Procederemos a introducir la clave de edicion Enterprise.
		slmgr /ipk YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 4 (
		goto :VERMENWIN
	)
:: Apartir de esta parte comienza el menu de Windows Server 2008 R2
:VERMENUWINDOWSSERVER2008R2
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows Server 2008 R2           ::
	echo   :: Eleja su Edicion de Windows Server 2008 R2     ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Datacenter                        ::
	echo   :: Presione 2 - Standard                          ::
	echo   :: Presione 3 - Enterprise                        ::
	echo   :: Presione 4 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234 >nul
	SET VERSIONWINSERVER=%ERRORLEVEL%
	IF %VERSIONWINSERVER% EQU 1 (
		Echo Procederemos a introducir la clave de edicion Datacenter.
		slmgr /ipk 74YFP-3QFB3-KQT8W-PMXWJ-7M648
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 2 (
		Echo Procederemos a introducir la clave de edicion Standard.
		slmgr /ipk YC6KT-GKW9T-YTKYR-T4X34-R7VHC
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 3 (
		Echo Procederemos a introducir la clave de edicion Enterprise.
		slmgr /ipk 489J6-VHDMP-X63PK-3K798-CPX3Y
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 4 (
		goto :VERMENWIN
	)
:: Apartir de esta parte comienza el menu de Windows Server 2012
:VERMENUWINDOWSSERVER2012
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows Server 2012              ::
	echo   :: Eleja su Edicion de Windows Server 2012        ::
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
		Echo Procederemos a introducir la clave de edicion Datacenter.
		slmgr /ipk 48HP8-DN98B-MYWDG-T2DCC-8W83P
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 2 (
		Echo Procederemos a introducir la clave de edicion Standard.
		slmgr /ipk XC9B7-NBPP2-83J2H-RHMBY-92BT4
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 3 (
		Echo Procederemos a introducir la clave de edicion Essentials.
		slmgr /ipk HTDQM-NBMMG-KGYDT-2DTKT-J2MPV
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 4 (
		goto :VERMENWIN
	)
:: Apartir de esta parte comienza el menu de Windows Server 2012 R2
:VERMENUWINDOWSSERVER2012R2
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows Server 2012 R2           ::
	echo   :: Eleja su Edicion de Windows Server 2012 R2     ::
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
		Echo Procederemos a introducir la clave de edicion Datacenter.
		slmgr /ipk W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 2 (
		Echo Procederemos a introducir la clave de edicion Standard.
		slmgr /ipk D2N9P-3P6X9-2R39C-7RTCD-MDVJX
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 3 (
		Echo Procederemos a introducir la clave de edicion Essentials.
		slmgr /ipk KNC87-3J2TX-XB4WP-VCPJV-M4FWM
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 4 (
		goto :VERMENWIN
	)
:: Apartir de esta parte comienza el menu de Windows Server 2016
:VERMENUWINDOWSSERVER2016
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows Server 2016              ::
	echo   :: Eleja su Edicion de Windows Server 2016        ::
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
		Echo Procederemos a introducir la clave de edicion Datacenter.
		slmgr /ipk CB7KF-BWN84-R7R2Y-793K2-8XDDG
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 2 (
		Echo Procederemos a introducir la clave de edicion Standard.
		slmgr /ipk WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 3 (
		Echo Procederemos a introducir la clave de edicion Essentials.
		slmgr /ipk JCKRF-N37P4-C2D82-9YXRT-4M63B
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 4 (
		goto :VERMENWIN
	)
::Apartir de esta parte comienza el menu de Windows Server 2019
:VERMENUWINDOWSSERVER2019
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows Server 2019              ::
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
		Echo Procederemos a introducir la clave de edicion Datacenter.
		slmgr /ipk WMDGN-G9PQG-XVVXX-R3X43-63DFG
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 2 (
		Echo Procederemos a introducir la clave de edicion Standard.
		slmgr /ipk N69G4-B89J2-4G8F4-WWYCC-J464C
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 3 (
		Echo Procederemos a introducir la clave de edicion Essentials.
		slmgr /ipk WVDHN-86M7X-466P6-VHXV7-YY726
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 4 (
		goto :VERMENWIN
	)
::Apartir de esta parte comienza el menu de Windows Server 2022
:VERMENUWINDOWSSERVER2022
	cls>nul 
	cls
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Activacion de Windows Server 2022              ::
	echo   :: Eleja su Edicion de Windows Server 2022        ::
	echo   :: Sub Menu                                       ::
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	echo   :: Presione 1 - Datacenter                        ::
	echo   :: Presione 2 - Datacenter Azure Edition          ::
	echo   :: Presione 3 - Standard                          ::
	echo   :: Presione 4 - Volver al Menu anterior           ::	
	echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
	choice /c 1234 >nul
	SET VERSIONWINSERVER=%ERRORLEVEL%
	IF %VERSIONWINSERVER% EQU 1 (
		Echo Procederemos a introducir la clave de edicion Datacenter.
		slmgr /ipk WX4NM-KYWYW-QJJR4-XV3QB-6VM33
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 2 (
		Echo Procederemos a introducir la clave de edicion Datacenter Azure Edition.
		slmgr /ipk NNTBV8-9K7Q8-V27C6-M2BTV-KHMXV
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 3 (
		Echo Procederemos a introducir la clave de edicion Standard.
		slmgr /ipk VDYBN-27WPP-V4HQT-9VMD4-VMK7H
		goto :ACTIVATIONWINDOWS
	)
	IF %VERSIONWINSERVER% EQU 4 (
		goto :VERMENWIN
	)
::Apartir de esta parte comienza los menus de Office
:OFFICE
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Activacion de Office 2013, 2016, 2019 y 2021   ::
echo   :: Sub Menu                                       ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - Elija su paquete de Office 2013   ::
echo   :: Presione 2 - Elija su paquete de Office 2016   ::
echo   :: Presione 3 - Elija su paquete de Office 2019   ::
echo   :: Presione 4 - Elija su paquete de Office 2021   ::
echo   :: Presione 5 - Volver al Menu Principal          ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 12345 >nul
SET OFFICEOPTION=%ERRORLEVEL%
IF %OFFICEOPTION% EQU 1 (
	goto :VERMENUOFFICE13
)
IF %OFFICEOPTION% EQU 2 (
	goto :VERMENUOFFICE16
)
IF %OFFICEOPTION% EQU 3 (
	goto :VERMENUOFFICE19
)
IF %OFFICEOPTION% EQU 4 (
	goto :VERMENUOFFICE21
)
IF %OFFICEOPTION% EQU 5 (
	goto :Principal
)
::Apartir de esta parte comienza el menu de Office 2013
:VERMENUOFFICE13
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Elija su paquete de Office 2013                ::
echo   :: Sub Menu                                       ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - ProPlus 2013                      ::
echo   :: Presione 2 - Project 2013                      ::
echo   :: Presione 3 - Visio 2013                        ::
echo   :: Presione 4 - Voler al Menu anterior            ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 1234 >nul
SET EDITOFFICEOPTION=%ERRORLEVEL%
IF %EDITOFFICEOPTION% EQU 1 (
	Echo Procederemos a introducir la clave de ProPlus 2013.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15")&(if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14")
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:GVGXT >nul&cscript //nologo ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 2 (
	Echo Procederemos a introducir la clave de Project 2013.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15")&(if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14")
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:2342K >nul&cscript //nologo ospp.vbs /inpkey:FN8TT-7WMH6-2D4X9-M337T-2342K >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 3 (
	Echo Procederemos a introducir la clave de Visio 2013.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15")&(if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14")
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:RM3B3 >nul&cscript //nologo ospp.vbs /inpkey:C2FG9-N6J68-H8BTJ-BW3QX-RM3B3 >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 4 (
	Echo Volviendo al Menu de Office 2013, 2016, 2019 y 2021
	pause
	goto :OFFICE
)
::Apartir de esta parte comienza el menu de Office 2016
:VERMENUOFFICE16
cls>nul
cls
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Elija su paquete de Office 2016                ::
echo   :: Sub Menu                                       ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
echo   :: Presione 1 - ProPlus 2016                      ::
echo   :: Presione 2 - Project 2016                      ::
echo   :: Presione 3 - Visio 2016                        ::
echo   :: Presione 4 - Voler al Menu anterior            ::
echo   ::::::::::::::::::::::::::::::::::::::::::::::::::::
choice /c 1234 >nul
SET EDITOFFICEOPTION=%ERRORLEVEL%
IF %EDITOFFICEOPTION% EQU 1 (
	Echo Procederemos a introducir la clave de ProPlus 2016.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 2 (
	Echo Procederemos a introducir la clave de Project 2016.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:G83KT >nul&cscript //nologo ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 3 (
	Echo Procederemos a introducir la clave de Visio 2016.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:RJRJK >nul&cscript //nologo ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 4 (
	Echo Volviendo al Menu de Office 2013, 2016, 2019 y 2021
	pause
	goto :OFFICE
)
::Apartir de esta parte comienza el menu de Office 2019
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
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 2 (
	Echo Procederemos a introducir la clave de Project 2019.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:PKD2B >nul&cscript //nologo ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 3 (
	Echo Procederemos a introducir la clave de Visio 2019.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\VisioPro2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\VisioPro2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:7VCBB >nul&cscript //nologo ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 4 (
	Echo Volviendo al Menu de Office 2013, 2016, 2019 y 2021
	pause
	goto :OFFICE
)
::Apartir de esta parte comienza el menu de Office 2021
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
	Echo Procederemos a introducir la clave de ProPlus 2021.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6F7TH >nul&cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 2 (
	Echo Procederemos a introducir la clave de Project 2021.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProjectPro2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:QV9H8 >nul&cscript //nologo ospp.vbs /inpkey:KFTNWT-C6WBT-8HMGF-K9PRX-QV9H8 >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 3 (
	Echo Procederemos a introducir la clave de Visio 2021.
	echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\VisioPro2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\VisioPro2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
	echo.
	echo ============================================================================
	echo Activating your Office...
	cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:K2HT4 >nul&cscript //nologo ospp.vbs /inpkey:KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4 >nul&set i=1
	goto :ACTIVATIONOFFICE
)
IF %EDITOFFICEOPTION% EQU 4 (
	Echo Volviendo al Menu de Office 2013, 2016, 2019 y 2021
	pause
	goto :OFFICE
)
::Apartir de esta parte comienza el menu para cargar el Servidor KMS y realizar la Activacion
:ACTIVATIONOFFICE
if %i%==1 set KMS_Sev=s8.uk.to
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo ============================================================================&choice /n /c YN /m "Desea realizar otra operación [Y,N]?" & if errorlevel 2 exit) || (echo La conexion al servidor KMS fallo! Intentando conectarme a otro ... & echo Por favor espera... & echo. & echo. & set /a i+=1 & goto ACTIVATIONOFFICE21)
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
