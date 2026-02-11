# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NXCustomViewDxfExporter is a Siemens NX automation plugin (class library) written in C# targeting .NET Framework 4.8. It uses the NXOpen API from Siemens NX 2406 and compiles to a DLL that is loaded and executed within the NX CAD environment. It exports custom views from the active part as individual DXF files.

## Build

Build using MSBuild (Visual Studio 18/Community). The primary configuration is **x64 Debug**:

```
msbuild NXCustomViewDxfExporter\NXCustomViewDxfExporter.csproj /p:Configuration=Debug /p:Platform=x64
```

Output DLL: `NXCustomViewDxfExporter\bin\x64\Debug\NXCustomViewDxfExporter.dll`

## Dependencies

All NXOpen assemblies are referenced from `C:\Program Files\Siemens\NX2406\NXBIN\managed\`:
- NXOpen.dll, NXOpen.UF.dll, NXOpen.Utilities.dll, NXOpenUI.dll

These are not bundled with the project; Siemens NX 2406 must be installed locally.

## Architecture

- **Main source**: `NXCustomViewDxfExporter\CustomViewDxfExporter.cs`
- **Entry point**: `CustomViewDxfExporter.Main(string[] args)` — standard NXOpen journal entry point invoked when NX loads the DLL
- **GetUnloadOption**: Required NXOpen callback that controls library lifecycle (set to unload immediately after execution)
- The plugin runs inside the NX process, accessing the active session via `Session.GetSession()`

## Running

The DLL is executed from within Siemens NX: **File > Execute > NX Open** or via the NX journal/macro system. It cannot be run standalone.
