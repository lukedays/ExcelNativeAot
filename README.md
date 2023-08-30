## Excel Native AOT

Proof-of-concept of using the newly announced .NET Native AOT (Ahead-of-time) compilation features to build Excel XLL add-ins.
The work is preliminary but very promising.

### Features

- 100% C# with unmanaged/native methods directly callable by the Excel C API
- No need for .NET runtimes for deployment
- Better performance (possibly?)

### Instructions

- Clone the project
- Debug > Start Debugging by using the `Excel` settings
- Debug > Attach to Process... and choose `EXCEL.EXE`
- Open `TestSheet.xlsx`
- Tweak settings/paths if necessary

### Test machine

- Visual Studio 2022 17.8.0 Preview 1.0
- .NET 8.0.100-preview.7.23376.3
- Microsoft® Excel® 365 (Version 2307) 64-bit 
- Windows 11 Home 22H2

### Todo

- Add all possible XlOper types
- Use code generation to wrap marshalling of variables and method generation
- Integrate into an existing project (candidates? Govert...?) or create a separate NuGet
- Find a way to auto-attach to the Excel process...

### Credits

- [Excel-DNA](https://excel-dna.net/) / [@govert](https://github.com/govert)
- [xlOil](https://xloil.readthedocs.io/en/stable/Introduction.html) / [@cunnane](https://github.com/cunnane)
- [Native AOT samples](https://github.com/dotnet/samples/tree/main/core/nativeaot)
- [Tutorial for building XLL's starting by the SDK](https://github.com/asavine/xlCppTutorial)
- Hundreds of StackOverflow pages searching for random C/C++/DllImport/marshalling errors
