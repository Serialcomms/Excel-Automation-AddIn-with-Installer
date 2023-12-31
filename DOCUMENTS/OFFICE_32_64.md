
```
  32/64 bit registration is driven by each Installer Project's Custom Actions

  Separate projects are provided and configured for 32-bit and 64-bit builds.
   
  Change Property setting Run64Bit to False for 32-bit Office registration

  Right-hand click on each category to change property as per screenshot below
  
  Install   > Primary Output > Properties > Run64Bit = False
  Commit    > Primary Output > Properties > Run64Bit = False
  Rollback  > Primary Output > Properties > Run64Bit = False
  UnInstall > Primary Output > Properties > Run64Bit = False

  Settings are then configured automatically by the installer class

  Environment.SystemDirectory = C:\Windows\SysWow64 for 32-bit install
  Environment.SystemDirectory = C:\Windows\System32 for 64-bit install

  Registry = HKEY_CLASSES_ROOT\WOW6432Node\CLSID\.. for 32-bit install
  Registry = HKEY_CLASSES_ROOT\CLSID\..             for 64-bit install

```

<img src="/SCREENSHOTS/CUSTOM_ACTIONS_AUTO_INSTALLER.png" alt="Properties" title="Project Properties" width="75%" height="75%">  

