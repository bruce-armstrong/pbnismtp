@echo off
cls
SETLOCAL

::*****************************
::* set environment variables *
::*****************************

::*
::* default location for VisualStudio.NET can be found in the registry key:
::*
::* HKLM\SOFTWARE\Microsoft\VisualStudio\8.0\Setup\VS\ProductDir
::*
::* default location for VisualC++.NET can be found in the registry key:
::*
::* HKLM\SOFTWARE\Microsoft\VisualStudio\8.0\Setup\VC\ProductDir
::*
::* change the values of the VS_INSTALL_DIR and VC8_DIR environment variables
::* below as necessary
::*

::set VS_INSTALL_DIR="C:\Program Files\Microsoft Visual Studio 8"
set VS_INSTALL_DIR="C:\Program Files\Microsoft Visual Studio 8"
set VC8_DIR=%VS_INSTALL_DIR%\Vc
set TARGET_DIR=%VC8_DIR%\VCWizards\PBNIWizard

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

copy .\VCWizards\PBNIWizard8\1033\*.* %TARGET_DIR%\1033\
copy .\VCWizards\PBNIWizard8\HTML\1033\*.* %TARGET_DIR%\HTML\1033\
copy .\VCWizards\PBNIWizard8\Images\*.* %TARGET_DIR%\Images\
copy .\VCWizards\PBNIWizard8\Scripts\1033\*.* %TARGET_DIR%\Scripts\1033\
copy .\VCWizards\PBNIWizard8\Templates\1033\*.* %TARGET_DIR%\Templates\1033\

::***************************
::* register the new wizard *
::***************************

copy .\VCProjects 8.0\PBNIWizard.vsdir %VC8_DIR%\vcprojects\
copy .\VCProjects 8.0\PBNIWizard.vsz %VC8_DIR%\vcprojects\
copy .\VCProjects 8.0\PBNIWizard.ico %VC8_DIR%\vcprojects\


:end
ENDLOCAL