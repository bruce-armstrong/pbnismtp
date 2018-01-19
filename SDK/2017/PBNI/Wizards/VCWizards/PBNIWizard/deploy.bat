@echo off
cls
SETLOCAL

::*****************************
::* set environment variables *
::*****************************

::*
::* default location for VisualStudio.NET can be found in the registry key:
::*
::* HKLM\SOFTWARE\Microsoft\VisualStudio\7.1\Setup\VS\ProductDir
::*
::* default location for VisualC++.NET can be found in the registry key:
::*
::* HKLM\SOFTWARE\Microsoft\VisualStudio\7.1\Setup\VC\ProductDir
::*
::* change the values of the VS_INSTALL_DIR and VC7_DIR environment variables
::* below as necessary
::*

::set VS_INSTALL_DIR="C:\Program Files\Microsoft Visual Studio .NET"
set VS_INSTALL_DIR="C:\Microsoft\Visual Studio .NET 2003"
set VC7_DIR=%VS_INSTALL_DIR%\Vc7
set TARGET_DIR=%VC7_DIR%\VCWizards\PBNIWizard

::*****************************
::* create target directories *
::*****************************

mkdir %TARGET_DIR%
mkdir %TARGET_DIR%\1033
mkdir %TARGET_DIR%\HTML\1033
mkdir %TARGET_DIR%\Images
mkdir %TARGET_DIR%\Scripts\1033
mkdir %TARGET_DIR%\Templates\1033

::************************************
::* copy files to target directories *
::************************************

copy .\VCWizards\PBNIWizard\1033\*.* %TARGET_DIR%\1033\
copy .\VCWizards\PBNIWizard\HTML\1033\*.* %TARGET_DIR%\HTML\1033\
copy .\VCWizards\PBNIWizard\Images\*.* %TARGET_DIR%\Images\
copy .\VCWizards\PBNIWizard\Scripts\1033\*.* %TARGET_DIR%\Scripts\1033\
copy .\VCWizards\PBNIWizard\Templates\1033\*.* %TARGET_DIR%\Templates\1033\

::***************************
::* register the new wizard *
::***************************

copy .\VCProjects 7.1\PBNIWizard.vsdir %VC7_DIR%\vcprojects\
copy .\VCProjects 7.1\PBNIWizard.vsz %VC7_DIR%\vcprojects\
copy .\VCProjects 7.1\PBNIWizard.ico %VC7_DIR%\vcprojects\


:end
ENDLOCAL