## Demonstration Build

The repository is intended as a demonstration / proof-of-concept build only.

It is _not_ intended for production use and distribution to end-users.

A new Production Build should be developed and tested before deployment.

*An 'Intermediate' level of experience in using Visual Studio is suggested to ensure a functional deliverable.*

## Production Build

<details><summary>Create new Visual Studio Solution & Projects</summary>
<p>

See [^1] for further information on Solutions and Projects.

<details><summary>Create new Visual Studio Solution</summary>
<p>

* Start Visual Studio and select `Create a New Project`.

* Select [`Blank Solution`](/SCREENSHOTS/VISUAL_STUDIO_NEW_BLANK_SOLUTION.png) as the Project Template and save with a name of your choice.

* In Solution Explorer, Right-Hand Click the above and select Add > New Project

</p>
</details> 

<details><summary>Add new Visual Studio .Net Project</summary>
<p>

<details><summary>Add new Project</summary>
<p>

* Add a new [Class Library .NET Framework](/SCREENSHOTS/VISUAL_STUDIO_NEW_CLASS_LIBRARY.png) Project and save with a name of your choice.
* In Solution Explorer, expand References and add [5 new entries as shown](/SCREENSHOTS/VISUAL_STUDIO_REFERENCES.png)
* Right-Hand click the new solution and select View Properties > Application.
* Check that `ASSEMBLY_NAME` and `ROOT NAMESPACE` are correct for your usage.
* Check that option [ ] `Make Assembly COM Visible` is not selected in Application > Assembly Information
* Check that option [ ]  `Register for COM interop` is not selected in Compile

</p>
</details> 

<details><summary>Add new COM Class</summary>
<p>
  
* Add a new [COM Class](/SCREENSHOTS/VISUAL_STUDIO_NEW_COM_CLASS.png) vb file to the Project and save with suggested name `Functions.vb`
  
  The new COM Class file will have new GUIDs created automatically which are valid for production use.
  
  Edit this file to add your User Defined Functions and change the general structure of it to resemble the demonstrator.
  
</p>
</details> 

<details><summary>Add new Partial Class</summary>
<p>

* Add a new [Class](/SCREENSHOTS/VISUAL_STUDIO_NEW_CLASS_DEFINITION.png) and save with suggested name `Interop.vb`

Replace the entire contents of the new file with the demonstrator version. 

Ensure that `Partial Public Class Functions` matches the Class Name of your main Functions class.

</p>
</details> 

<details><summary>Add new Installer Class</summary>
<p>
  
* Add a new [Installer Class](/SCREENSHOTS/VISUAL_STUDIO_NEW_INSTALLER_CLASS.png) and save with suggested name `Installer.vb`

Replace the entire contents of the new file with the demonstrator version.

Ensure that references to `Functions` in `Sub New()` match the Class Name of your main Functions class.

Build the project and check that it completes successfully before continuing. 

</p>
</details> 

</p>
</details> 

<details><summary>Create New Visual Studio Setup Projects</summary>
<p>

<details><summary>Add Setup Projects</summary>
<p>

In Solution Explorer, right-hand click the main Solution and

* Add a new [Setup Project](/SCREENSHOTS/VISUAL_STUDIO_NEW_SETUP_PROJECT.png) Project and save with a name of your choice for 32-Bit Install.

* Add a second new [Setup Project](/SCREENSHOTS/VISUAL_STUDIO_NEW_SETUP_PROJECT.png) Project and save with a name of your choice for 64-Bit Install.

</p>
</details> 

<details><summary>Configure Setup Projects</summary>
<p>

In Solution Explorer, right-hand click each Setup Project and 

1. Select Add > Project Output and add the Primary Output
2. Select [View > Custom Actions](/SCREENSHOTS/CUSTOM_ACTIONS_VIEW.png) and add the Primary Output to each of the [four categories shown](/SCREENSHOTS/CUSTOM_ACTIONS_AUTO_INSTALLER.png)
3. Press the F4 key and set Company Name etc. to values of your choice. 

Right-hand click the Primary Output in each of the four categories and 

1. Rename the Primary Output (optional)
2. Check that Property InstallerClass = True
3. Set Property [Run64Bit to True](/SCREENSHOTS/CUSTOM_ACTIONS_RUN64BIT.png) for 64-Bit Office and False for 32-Bit Office.

Note that the same Primary Output .dll file is used for both 32-bit and 64-bit installers. 

</p>
</details> 

</p>
</details> 

</p>
</details> 



[^1]:https://learn.microsoft.com/en-us/visualstudio/get-started/tutorial-projects-solutions?view=vs-2022

