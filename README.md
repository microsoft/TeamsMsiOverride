# Teams MSI Override
Microsoft Teams supports installation through an MSI installer, referred to as the Teams Machine-Wide Installer. This installer is used by Microsoft Office to install Teams or may be used by organizations installing Teams through a deployment package.

The Teams Machine-Wide Installer does not normally get updated, and per-user instances of Teams installed into a user's profile will normally not be affected by changes to the Teams-Machine-Wide Installer.

There may be cases where an organization needs to update the Teams Machine-Wide Installer or forcibly update the per-user instances of Teams, perhaps for a critical security release, or if the normal update process is failing, or because a machine is shared, and new users are getting an outdated Teams installation.

In the event an organization needs to update the Teams Machine-Wide Installer, they can use a feature named MSI Override to update the MSI installed on a machine and allow per-user instances of Teams to update from the MSI.

For additional technical details on the Teams Machine-Wide Installer and MSI Override, please see [Details](Details.md)

To implement MSI Override, you can use the PublishLatestVersion.ps1 and CheckMsiOverride.ps1 scripts.

[![Version](https://img.shields.io/github/v/release/microsoft/TeamsMsiOverride?label=latest%20version)](https://github.com/microsoft/TeamsMsiOverride/releases/latest/download/TeamsMsiOverride.zip)
[![Downloads](https://img.shields.io/github/downloads/microsoft/TeamsMsiOverride/total)](https://github.com/microsoft/TeamsMsiOverride/releases/latest/download/TeamsMsiOverride.zip)

## Getting Started
PowerShell **5.0** (or greater) must be installed on the host machine. Click [here](https://github.com/powershell/powershell) for details on how to get the latest version for your computer. 

### PublishLatestVersion
The PublishLatestVersion.ps1 script will retrieve the most recent Teams MSI installers and store them onto a file share.

To use the script, follow these steps:
1) [Download](https://github.com/microsoft/TeamsMsiOverride/releases/latest/download/TeamsMsiOverride.zip) the latest version of the script package.
2) Create a file share (i.e., \\\\server\TeamsUpdateShare) which is accessible by all users which require the Teams MSI Override deployment. Ensure general users have read access, and only a few users, such as your IT administrators, have write access (these are your MSI Override Administrators).
3) Copy the PublishLatestVersion.ps1 script to a location available to your MSI Override Administrators.
4) Run the PublishLatestVersion.ps1 script, providing the path to the file share as follows:

   ```PublishLatestVersion.ps1 -BaseShare \\server\TeamsUpdateShare```

At this point the file share should be populated with a new folder containing the latest Teams MSI installers, as well as a Version.txt file which indicates which version is the latest.
Going forward the MSI Override Administrators can run the PublishLatestVersion.ps1 script to retrieve the latest Teams MSI installers at any time.

### CheckMsiOverride
The CheckMsiOverride script is intended to run on user machines, and will set the required registry key, and update the Teams Machine-Wide Installer to the most recent version from the file share.

It must be run with Administrative privileges. It may be ideal to have it run from the SYSTEM account.

The CheckMsiOverride.ps1 script can be deployed in various ways; we will provide an example here using Scheduled Tasks. 

The script is signed, so the user executing the script on each machine will require an execution policy of at least RemoteSigned.

#### Scheduled Task
To deploy this script as a Scheduled Task you can use the following steps:
1) [Download](https://github.com/microsoft/TeamsMsiOverride/releases/latest/download/TeamsMsiOverride.zip) the latest version of the script package.
2) Copy the CheckMsiOverride.ps1 script to the file share you created (\\\\server\TeamsUpdateShare\CheckMsiOverrride.ps1), or any other equally accessible location per your organizations policies.
3) Create a scheduled task which executes the script as follows:

   ```powershell.exe -File \\share\TeamsUpdateShare\CheckMsiOverride.ps1 -BaseShare \\share\TeamsUpdateShare```
4) Specify a schedule that is appropriate for your organization. If no update is required, the script will make no changes, so there are no issues running it often (such as daily).

#### AllowInstallOvertopExisting
When installing the Teams Machine-Wide Installer originally, it could have been installed in 3 main ways, relative to the current user:
1) Installed by any user using ALLUSERS=1 parameter
2) Installed by the current user without the ALLUSERS=1 parameter
3) Installed by a different user without the ALLUSERS=1 parameter

For scenarios 1 and 2, the script can perform an in-place upgrade of the MSI.

For scenario 3 the current user is not "aware" that the MSI has been installed, and so it is not able to do an in-place upgrade. In this case the script will, by default, exit with an error.

If you pass the -AllowInstallOvertopExisting switch into the script, it will permit the script to instead perform an installation of the MSI for the current user. This will overwrite the existing files, allowing them to be updated to the correct version.

   ```powershell.exe -File \\share\TeamsUpdateShare\CheckMsiOverride.ps1 -BaseShare \\share\TeamsUpdateShare -AllowInstallOvertopExisting```

If this occurs, however, two different users will have separate installation entries created.  If either user uninstalls the Teams Machine-Wide Installer, the files will be removed, and it will be shown as uninstalled for the user performing the uninstall, but the second user will still show an installation entry present, even though the files have been removed.

#### OverwritePolicyKey
By default, the script will populate the AllowMsiOverride key only if it does not already exist. Therefore, if your organization wants to push a value of 0 to some users, this value will remain even if the script is ran.

If you want to forcibly reset the value back to 1 for all users, you can pass the -OverwritePolicyKey switch.

   ```powershell.exe -File \\share\TeamsUpdateShare\CheckMsiOverride.ps1 -BaseShare \\share\TeamsUpdateShare -OverwritePolicyKey```

#### Diagnostics
CheckMsiOverride.ps1 will save a trace file to %TEMP%\TeamsMsiOverrideTrace.txt

It will also write to the Application event log with the source "TeamsMsiOverride" for any failures, or if an update completed successfully.

# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

# Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
