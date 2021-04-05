## Teams Machine-Wide Installer
Microsoft Teams supports installation through an MSI installer, referred to as the Teams Machine-Wide Installer. This installer is used by Microsoft Office to install Teams, or may be used by organizations installing Teams through a deployment package.

When installed through this method, the MSI installer places an EXE installer onto the machine in the Program Files folder, and a [Run key](https://docs.microsoft.com/en-us/windows/win32/setupapi/run-and-runonce-registry-keys) named TeamsMachineInstaller is created in the Local Machine hive.

When a user logs into the machine, this [Run key](https://docs.microsoft.com/en-us/windows/win32/setupapi/run-and-runonce-registry-keys) will execute, installing Microsoft Teams into the per-user profile location. From that point the per-user instance of Teams should automatically update itself.

The Teams Machine-Wide Installer does not update itself, so the installer present on a given machine will generally remain at the version first installed. 
Since this is just used to initially install Teams, and then the per-user profile instance of Teams will automatically update itself, this is generally not an issue, except in the case of shared computers where new users are logging into them frequently.

Even if the Teams Machine-Wide Installer is updated, it will normally not affect the per-user instance of Teams installed into a user's profile.

## Teams MSI Override

For Teams MSI Override to work, two things must occur:
1) The installed Teams Machine-Wide Installer (MSI) must be updated to the target version
2) A DWORD registry key must be created:

   ```HKLM\Software\Policies\Microsoft\Office\16.0\Teams\AllowMsiOverride = 1```

The next time the user signs into Windows and the TeamsMachineInstaller [Run key](https://docs.microsoft.com/en-us/windows/win32/setupapi/run-and-runonce-registry-keys) is executed, with AllowMsiOverride set to 1, it will check if a newer version of Teams is available from the Teams Machine-Wide Installer and install it into the per-user profile instance.