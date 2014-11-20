[:arrow_forward:](Installation.md)

# Building qXL

- [System requirements](Building-qXL.md#system-requirements)
- [Build steps](Building-qXL.md#build-steps)
- [Troubleshooting](Installation.md#troubleshooting)

## System requirements

The following are required to build `qXL`:
- [Microsoft .NET Framework 4](http://www.microsoft.com/en-us/download/details.aspx?id=17718)
- [WiX Toolset v3.8](https://wix.codeplex.com/)
- `AssemblyInfo` task from [MSBuild Community Tasks](https://github.com/loresoft/msbuildtasks)

The `qXL` solution is compatible with Visual Studio 2010.

> Note:
> 
> Visual Studio Express does not support extensions thus the `ExcelAddInDeploy` project (WIX installer) cannot be opened.

## Build steps

Make sure that the `msbuild.exe` is on system `PATH`. You can extend `PATH` with following command:
```shell
set PATH=%PATH%;C:\Windows\Microsoft.NET\Framework64\v4.0.30319
```

To fully rebuild the solution and create the installer, execute:
```shell
msbuild qXL.sln /t:Rebuild /p:Configuration=Release
```

The installer is then built to the `ExcelAddInDeploy\bin\Release` folder.

> Note:
> 
> Make sure that installer is not built in the `Debug` configuration.

## Version stamping

The version number in AssemblyInfo and installer is generated upon build based on environmental variables:
 - `VERSION_MAJOR`,
 - `VERSION_MINOR`,
 - `VERSION_REVISION`,
 - `VERSION_BUILD`.
    
Version number follows the pattern `VERSION_MAJOR.VERSION_MINOR.VERSION_REVISION.VERSION_BUILD`, where each version element is represented as 16 bit integer. 

In addition the `qXL` name can be post-fixed if the `VERSION_TYPE`. This can be used to create `BETA` and `RC` flavours of build.


## Troubleshooting

If a build is not automatically fetching the dependencies, then issuing the command 

```shell
nuget install packages.config -o packages
```

on the command line in the `qXL` root directory could be used to manually fetch those dependencies.
