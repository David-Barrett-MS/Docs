# Overview

This guide describes the process to create a Setup project (.msi installer) for a VSTO Outlook add-in developed using Visual Studio 2019.  Add-ins created with earlier versions of Visual Studio may need their existing Setup project deleted and recreated.

# Install the Visual Studio Installer Projects extension

By default, Visual Studio 2019 does not have a Setup project template.  These templates can be added by installing the Visual Studio Installer Projects extension.

* Open Visual Studio.

* Select Extensions... Manage Extensions.

* Search online for Microsoft Visual Studio Installer
![Search for Microsoft Visual Studio 2022 Installer](images/Visual%20Studio%202022%20Installer%20module.png)

* Download and install the extension, and then restart Visual Studio.

# Add and configure a Setup project to the Outlook add-in solution

* Open the Outlook add-in solution.

* Add new project to the solution (File... Add... New Project...).

* Select Setup project as the type, and then complete the subsequent steps in the wizard to add the project to the solution.
![Add Setup project to solution](../VSTO%20Setup%20with%20Visual%20Studio%202019/images/Add%20new%20Setup%20project.png)

* The default view for the Setup project should be the File System view.  The first step is to add the primary output of the VSTO add-in project to the Application Folder.
![Add primary output of add-in project](../VSTO%20Setup%20with%20Visual%20Studio%202019/images/File%20System%20-%20add%20primary%20output.png)

* You also need to add the add-in's manifest and VSTO files from the add-in output directory.
![Add VSTO and manifest files](../VSTO%20Setup%20with%20Visual%20Studio%202019/images/File%20System%20-%20add%20manifest%20and%20VSTO.png)

* Next, the registry needs to be configured.  Open the registry view.
![Open Registry view](../VSTO%20Setup%20with%20Visual%20Studio%202019/images/Installer%20views.png)

* The final step is to configure the registry.  You should be able to copy the settings from the developer machine, remembering to update the manifest reference.  This installation will only work for the user that installs as the registry keys target HKCU.  For all-user add-ins, the appropriate registry keys need to be added to HKLM.  Further information on registry settings can be found here: https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins?view=vs-2019.  Useful information regarding 32 bit and 64 bit add-ins can be found here: https://learn.microsoft.com/en-us/previous-versions/troubleshoot/outlook/addins-are-registered-under-wow6432node.
![Add registry keys](../VSTO%20Setup%20with%20Visual%20Studio%202019/images/Registry%20-%20configure%20HKCU.png)

* The Setup project can now be built, and once done you'll find Setup.exe and an .msi in the output folder of the project.
