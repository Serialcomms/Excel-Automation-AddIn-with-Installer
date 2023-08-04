# Excel Automation Add-In with Installer in VB.Net
Excel User-Defined-Functions (UDF) Automation Add-In [^2] with integrated Installer, both written in VB.Net

Full Microsoft Visual Studio 2022 solution and project source code included in repository.

<details><summary>Background Information</summary>
<p>
  
Excel User-Defined Functions (UDFs) as developed in VB.Net have been around for many years, early examples of which are published here - 

http://www.cpearson.com/Excel/CreatingNETFunctionLib.aspx

https://www.codeproject.com/Articles/7753/Create-an-Automation-Add-In-for-Excel-using-NET


Whilst these functions work well, the deployment of them can be more problematic, particularly where end users may not be familiar with, or are permitted to run command-line utilities such as RegAsm.exe to complete the installation. 

</p>
</details>  

<details><summary>Design Goals</summary>
<p>

The design goals for this project are therefore :-

1.  Working Excel Automation Add-In with sample functions provided
2.  Integrated 'Click-Through' Installer, more familiar to end-users
3.  All development in VB.Net, using Microsoft Visual Studio 2022
4.  No third-party libraries or utilities required
5.  Coding style to support infrequent developers
6.  Configurable for 32-Bit or 64-Bit Office - see later for details

</p>
</details> 

<details><summary>Dependencies</summary>
<p>

A Windows PC with the following software installed is required to build the solution 

1.  Microsoft Windows 10, 64-Bit with .Net
2.  Microsoft Office/Excel 32-Bit or 64-Bit
3.  Microsoft Visual Studio 2022 (any edition)

A 'fresh build' of all the above components is recommended, on a dedicated development PC if possible, and with all updates applied.

Visual Studio should have the following items installed

* [.Net Desktop Development workload](/SCREENSHOTS/VISUAL_STUDIO_WORKLOAD_DOTNET_DESKTOP.png)
* [Office/Sharepoint Development workload](/SCREENSHOTS/VISUAL_STUDIO_WORKLOAD_OFFICE_DEVELOPMENT.png)
* [Visual Studio Installer Projects 2022 Extension](/SCREENSHOTS/VISUAL_STUDIO_EXTENSIONS.png)


</p>
</details> 

<details><summary>Optional Utilities</summary>
<p>

The following utility is useful to inspect the Registration process, but is not mandatory.

1. https://www.nirsoft.net/utils/registered_dll_view.html

</p>
</details> 

<details><summary>32/64 Bit Office</summary>
  
<p>

The Automation AddIn needs to be registered during the installation process. 
  
Different values need to be written to the Registry for 32-Bit and 64-Bit version of Office.

The installer class provides these values, but needs to be [configured correctly](/SCREENSHOTS/CUSTOM_ACTIONS_RUN64BIT.png) for the version required.

Separate installers should be built for each version required. 

A Universal 32/64 installer is not supported at this time, but could be developed.

</p>
</details> 

<details><summary>Automation Add-In - User Installation</summary>
  
<p>

Visual Studio generates two output files, `setup.exe` and `AUTO_INSTALLER.msi`

Either of these files can be distributed to, and run by end users, to install and uninstall as required. 

</p>
</details> 

<details><summary>Automation Add-In - Excel Configuration</summary>

<p>

After running the [installer](/SCREENSHOTS/USER_INSTALL_01.jpg), users need to configure Excel to enable the Automation Add-In.

From Excel > File > Options > [Add-Ins](/SCREENSHOTS/EXCEL_ADDIN_01.png) > [Manage Excel Add-Ins](/SCREENSHOTS/EXCEL_ADDIN_02.png) 

Click on Automation, scroll down and select [AUTOMATION.Functions](/SCREENSHOTS/EXCEL_ADDIN_03.png)

Click [OK](/SCREENSHOTS/EXCEL_ADDIN_04.png) to confirm

</p>
</details> 


<details><summary>Automation Add-In - Excel Formulas</summary>
<p>
  
Two sample [Excel formulas](/SCREENSHOTS/EXCEL_FORMULAS_01.png) are supplied

`=IFX()` in a Worksheet cell returns the text string `AUTO FX OK`

`=TIMENOW()` in a Worksheet cell returns the current time with milliseconds e.g. `12:34:56.789`

This is a 'Volatile' function and will re-calculate when the F9 key is pressed or another cell changes. 

Functions offered by the Add-In can be listed by clicking on Formulas > Insert Function and selecting [AUTOMATION.Functions](/SCREENSHOTS/EXCEL_INSERT_FUNCTION.png)
as a category


</p>
</details> 

<details><summary>Automation Add-In - Uninstalling</summary>
<p>

Users can uninstall the addin by Right-Clicking the Windows Start button and selecting [Apps and Features](/SCREENSHOTS/APPS_AND_FEATURES.png)

Scroll down to *Automation FX* and select Uninstall

</p>
</details> 

<details><summary>Implementation Notes</summary>
<p>

<details><summary>Installer Class Module</summary>
<p>

#### Installer Class Module
Class module `Installer.vb` performs the Assembly Registration and Registry updates required when the developer or end-user runs the installer .exe or .msi program. 

Tag `<System.ComponentModel.RunInstaller(True)>` is provided automatically by vb.net in file `Installer.Designer.vb` when a new Installer class module is added to a project.

This tag is used by the installer program to call `Public Overrides Sub Install(stateSaver As IDictionary)` via [Custom Action Properties](/SCREENSHOTS/CUSTOM_ACTIONS_INSTALLERCLASS.png) in module AUTO_INSTALLER.

Sub `Install` then calls `RegisterAssembly` which is functionally equivalent [^1] to running `RegAsm.exe` manually. 

`RegAsm.exe` itself  _uses methods exposed by RegistrationServices_ [^3]

</p>
</details> 

<details><summary>COM Configuration Options</summary>
<p>

#### COM Configuration Options 
The following points should therefore be observed to avoid performing any conflicting Registry updates during development and testing.

In module AUTO_FUNCTIONS > Properties, the options below should **not** be selected at any time.
- [ ] `Register for COM Interop` in section Compile 
- [ ] `Make assembly COM-Visible` in section Application > Assembly Information

Tags `<ComRegisterFunction>` and  `<ComUnRegisterFunction>` should also **not** be used in any module.

#### Primary Output Properties

In module AUTO_INSTALLER > Primary Output Properties, property Register should be set to vsdrpDoNotRegister

</p>
</details> 

</p>
</details> 

  
[^1]:https://learn.microsoft.com/en-us/dotnet/framework/interop/registering-assemblies-with-com

[^2]:https://support.microsoft.com/en-us/topic/excel-com-add-ins-and-automation-add-ins-91f5ff06-0c9c-b98e-06e9-3657964eec72

[^3]:https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.registrationservices?view=netframework-4.8.1
