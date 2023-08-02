# Excel Automation Add-In with Installer in VB.Net
Functional Excel Automation Add-In with Installer, both written in VB.Net

Full Microsoft Visual Studio 2022 solution and project source code.

<details><summary>Background Information</summary>
<p>
  
Excel User-Defined Functions (UDFs) as developed in VB.Net have been around for many years, early examples of which are published here - 

http://www.cpearson.com/Excel/CreatingNETFunctionLib.aspx

https://www.codeproject.com/Articles/7753/Create-an-Automation-Add-In-for-Excel-using-NET


Whilst these functions work well, the deployment of them can be more problematic, particularly where end users may not be familiar with, or permitted to, run command line utilities such as RegAsm to complete the installation. 

</p>
</details>  

<details><summary>Design Goals</summary>
<p>

The design goals for this project are therefore :-

1.  Working Excel Automation Add-In with sample functions
2.  Integrated 'Click-Through' Installer, more familiar to end-users
3.  All development in VB.Net, using Microsoft Visual Studio 2022
4.  No third-party libraries or utilities required
5.  Coding style to support infrequent developers
6.  Configurable for 32-Bit or 64-Bit Office - see later for details

</p>
</details> 

<details><summary>Dependencies</summary>
<p>

A PC with the following software installed is required. 

1.  Microsoft Windows 10, 64-Bit with .Net
2.  Microsoft Office/Excel 32-Bit or 64-Bit
3.  Microsoft Visual Studio 2022 (any edition)

A 'fresh build' of all the above components is recommended, on a dedicated development PC if possible, and with all updates applied.

</p>
</details> 

<details><summary>Optional Utilities</summary>
<p>

The following utility is useful to inspect the Registration process, but is not mandatory.

1. https://www.nirsoft.net/utils/registered_dll_view.html
  
2. 
</p>
</details> 

<details><summary>32/64 Bit Office</summary>
  
<p>

The Automation AddIn needs to be registered during the installation process. 
  
Different values need to be written to the Registry for 32-Bit and 64-Bit version of Office.

The installer provides these values, but needs to be configured correctly for the version required.

Separate installers should be built for each version required. 

A Universal 32/64 installer is not supported at this time, but could be developed.

</p>
</details> 

