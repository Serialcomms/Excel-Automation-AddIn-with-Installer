## Demonstration Build

The repository is intended as a demonstration / proof-of-concept build only.

It is not intended for production use, particularly where GUIDs as concerned.

Various changes are required to ensure it is suitable for distribution to end-users as a production build.

## Production Build

Two approaches are available to prepare a production build, in addition to changing the User-Defined Functions.

<details><summary>1. Edit existing Visual Studio Solution</summary>
<p>

* Search *ALL* solution and project file for GUIDs

* Change all GUIDs to newly-created values - (Tools - Create GUID)

* Check / Update `<ProgId("AUTOMATION.Functions")>` in Functions.vb

* Check / Update project AUTO_INSTALLER_nn properties - Press F4 to view

* Rebuild Solution and test fully before distribution

</p>
</details> 

<details><summary>2. Create new Visual Studio Solution</summary>
<p>

This is the preferred approach.

Start Visual Studio and select `Create a New Project`.

Select [`Blank Solution`](/SCREENSHOTS/VISUAL_STUDIO_NEW_BLANK_SOLUTION.png) as the Project Template.


See [^1] for further information on Solutions and Projects.

</p>
</details> 


[^1]:https://learn.microsoft.com/en-us/visualstudio/get-started/tutorial-projects-solutions?view=vs-2022

