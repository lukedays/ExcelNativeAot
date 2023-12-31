## Excel Native AOT

Proof-of-concept of using .NET Native AOT (ahead-of-time) compilation features to build Excel XLL add-ins.

### Features

- 100% C# with unmanaged/native methods directly callable by Excel (the C/C++ projects are used just for debugging)
- Access to the C/XLL (through `UnmanagedCallersOnly`) and COM APIs (through `GeneratedComInterface`)
- No need for .NET runtimes for deployment
- Better performance and/or hot reloading (possibly?)

### Instructions

- Clone the project
- Set `ExcelNativeDebugger` as the starting project if it's not already
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

### Credits

- [Excel-DNA](https://excel-dna.net/) / [@govert](https://github.com/govert)
- [xlOil](https://xloil.readthedocs.io/en/stable/Introduction.html) / [@cunnane](https://github.com/cunnane)
- [Native AOT samples](https://github.com/dotnet/samples/tree/main/core/nativeaot)
- [Tutorial for building XLL's starting by the SDK](https://github.com/asavine/xlCppTutorial)
- Bing Chat, StackOverflow, etc
