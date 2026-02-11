# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NXOpenTest is a Siemens NX automation plugin (class library) written in C# targeting .NET Framework 4.8. It uses the NXOpen API from Siemens NX 2406 and compiles to a DLL that is loaded and executed within the NX CAD environment.

## Build

Build using MSBuild (Visual Studio 2022). The primary configuration is **x64 Debug**:

```
msbuild NXOpenTest\NXOpenTest.csproj /p:Configuration=Debug /p:Platform=x64
```

Output DLL: `NXOpenTest\bin\x64\Debug\NXOpenTest.dll`

## Dependencies

All NXOpen assemblies are referenced from `C:\Program Files\Siemens\NX2406\NXBIN\managed\`:
- NXOpen.dll, NXOpen.UF.dll, NXOpen.Utilities.dll, NXOpenUI.dll

These are not bundled with the project; Siemens NX 2406 must be installed locally.

## Architecture

- **Entry point**: `Class1.Main(string[] args)` — standard NXOpen journal entry point invoked when NX loads the DLL
- **GetUnloadOption**: Required NXOpen callback that controls library lifecycle (set to unload immediately after execution)
- The plugin runs inside the NX process, accessing the active session via `Session.GetSession()`

## Running

The DLL is executed from within Siemens NX: **File > Execute > NX Open** or via the NX journal/macro system. It cannot be run standalone.
