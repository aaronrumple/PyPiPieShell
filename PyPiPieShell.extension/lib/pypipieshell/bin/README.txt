Optional: put ICSharpCode.AvalonEdit.dll here.

PyPiPie Shell uses the AvalonEdit editor control. pyRevit 7 ships it beside
its IronPython engines and the shell finds it there automatically. A copy in
this folder is used first, which helps with older pyRevit versions or
unusual installs.

Use the build that matches the .NET Revit runs on:
  Revit 2024 and older  -> net462 build (pyRevit bin\netfx\...)
  Revit 2025 and newer  -> net6.0/net8.0 build (pyRevit bin\netcore\...)

AvalonEdit is MIT licensed: https://github.com/icsharpcode/AvalonEdit
