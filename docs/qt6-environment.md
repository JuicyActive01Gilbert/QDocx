# Qt6 Environment

QDocx is a Windows-only Qt 6 project because it uses ActiveQt / `QAxObject` to call Office Word or WPS through COM automation.

## Required Components

- Windows
- Qt 6
- Qt modules: `Core`, `Gui`, `Widgets`, `AxContainer`
- CMake 3.21+
- MSVC C++ toolchain
- Office Word or WPS installed and registered for COM automation

## Configure

Replace `D:/Qt/6.x.x/msvcxxxx_64` with your local Qt 6 kit path.

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64 `
  -DCMAKE_PREFIX_PATH="D:/Qt/6.x.x/msvcxxxx_64"
```

## Build

```powershell
cmake --build build --config Release
```

## Test

```powershell
ctest --test-dir build -C Release --output-on-failure
```
