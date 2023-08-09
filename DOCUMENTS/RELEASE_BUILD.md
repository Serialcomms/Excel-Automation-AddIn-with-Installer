## Demonstration Build

The repository is intended as a demonstration / proof-of-concept build only.

It is _not_ intended for production use, specifically as new GUIDs as required.

Various changes are required to ensure it is suitable for distribution to end-users as a production build.

*An 'Intermediate' level of experience in using Visual Studio is suggested to ensure a functional deliverable.*

## Production Build

Two approaches are available to prepare a production build, in addition to changing the User-Defined Functions.

<details><summary>1. Edit existing Visual Studio Solution</summary>
<p>

* Search *ALL* solution and project file for GUIDs

* Change all GUIDs to newly-created values - (Tools > Create GUID)

* Check / Update `<ProgId("AUTOMATION.Functions")>` in Functions.vb

* Check / Update project `AUTO_INSTALLER_nn` properties - Press F4 to view

* Rebuild Solution and test fully before distribution to end-users.

</p>
</details> 

<details><summary>2. Create new Visual Studio Solution & Projects</summary>
<p>

_This is the preferred approach and should result in a 'cleaner' build with less errors._

See [^1] for further information on Solutions and Projects.

<details><summary>Create New Visual Studio Solution</summary>
<p>

* Start Visual Studio and select `Create a New Project`.

* Select [`Blank Solution`](/SCREENSHOTS/VISUAL_STUDIO_NEW_BLANK_SOLUTION.png) as the Project Template and save with a name of your choice.

* In Solution Explorer, Right-Hand Click the above and select Add > New Project

</p>
</details> 

<details><summary>Add New Visual Studio .Net Project</summary>
<p>

* Add a new [Class Library .NET Framework](/SCREENSHOTS/VISUAL_STUDIO_NEW_CLASS_LIBRARY.png) Project and save with a name of your choice.
* In Solution Explorer, expand References and add [5 new entries as shown](/SCREENSHOTS/VISUAL_STUDIO_REFERENCES.png)
* Right-Hand click the new solution and select View Properties > Application.
* Check that `ASSEMBLY_NAME` and `ROOT NAMESPACE` are correct for your usage.
* Check that [] `Make Assembly COM Visible` is not selected in Application > Assembly Information
* Check that [] `Register for COM interop` is not selected in Compile

<details><summary>Add new COM Class</summary>
<p>
  
* Add a new [COM Class](/SCREENSHOTS/VISUAL_STUDIO_NEW_COM_CLASS.png) vb file to the Project and save with suggested name `Functions.vb`
  
  The new COM Class file will have new GUIDs created automatically which are valid for production use.
  
  Edit this file to add your User Defined Functions and change the general structure to resemble the demonstrator.
  
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

</p>
</details> 

</p>
</details> 

<details><summary>Create New Visual Studio Installer Projects</summary>
<p>

In Solution Explorer, Right-Hand Click

* Add a new [Setup Project](/SCREENSHOTS/VISUAL_STUDIO_NEW_SETUP_PROJECT.png) Project and save with a name of your choice for 32-Bit Install.

* Add a second new [Setup Project](/SCREENSHOTS/VISUAL_STUDIO_NEW_SETUP_PROJECT.png) Project and save with a name of your choice for 64-Bit Install.

</p>
</details> 

</p>
</details> 





[^1]:https://learn.microsoft.com/en-us/visualstudio/get-started/tutorial-projects-solutions?view=vs-2022

