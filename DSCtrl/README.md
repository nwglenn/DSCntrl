# DSCtrl - Deskside Control Panel

## Introduction
The DSCtrl application is a set of Powershell scripts which run behind a Graphical User Interface (GUI). These scripts, in conjunction with the user interface aim to assist University Technology Office staff with day-to-day management tasks within the M.UTOSPA OU structure.

## Istallation
UTO Deskside staff can install DSCtrl via the package available in Software Center, or by extracting the latest source package anywhere on their system.

_IMPORTANT:_ The DSCtrl relies heavilly on functions and libraries provided by the Microsoft Remote Server Administration Tools (RSAT) package. Before using DSCtrl, please ensure that you have the latest version of RSAT installed on your system. UTO Deskside techs can install RSAT from Software Center for convinience.

## Usage
The DSCtrl release package comes with a shortcut file that can be used to launch the main user interface. Double-clicking the _DSCTRL.lnk_ file should be all that is necessary to run the application

In the event that you need to stat DSCtrl manually, you can use the following command from the Windows command line;

```
powershell.exe -STA -File <path to DSCtrl>\ds.ps1
```

## Troubleshooing
If you encounter any problems starting the program, try these troubleshooting steps

1. Check that your system is not blocking execution of local Powershell scripts.
2. Make sure that the latest version of RSAT is installed
3. Do not run DSCtrl from a networked drive
4. If using a version of Windows prior to Windows 10, install the .Net Management Framework 5.0 or later; UTOSPA machines should have this available in Software Center

## Log files
By default, DSCtrl logs all errors and status messages to ```%localappdata%\ASU\DSctrl\logs\```. The name of these logs follows this name format;

```
[function name]-[time stamp in NT Time].log
```

## License

This software is intended to be used by UTO Deskside staff and student workers. Please do not distribute it outside of the university.
(C) 2017 Arizona State University
