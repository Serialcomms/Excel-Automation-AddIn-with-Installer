## Demonstration Build

The repository is intended as a demonstration / proof-of-concept build only.

It is not intended for production use, particularly where GUIDs as concerned.

Various changes are required to ensure it is suitable for distribution to end-users as a production build.

## Production Build

Two approaches are available to prepare a production build, in addition to changing the User-Defined Functions.

<details><summary>Edit existing Visual Studio Solution</summary>
<p>

* Search *ALL* solution and project file for GUIDs

* Change all GUIDs to newly-created values.

</p>
</details> 

<details><summary>Create new Visual Studio Solution</summary>
<p>

This is the preferred approach.

Start Visual Studio and select `Create a New Project`.

Select [`Blank Solution`](/SCREENSHOTS/VISUAL_STUDIO_NEW_BLANK_SOLUTION.png)

https://learn.microsoft.com/en-us/visualstudio/get-started/tutorial-projects-solutions?view=vs-2022


</p>
</details> 




