# Windows Management Instrumentation (WMI) C++ Library

A modern C++ wrapper for Windows Management Instrumentation (WMI) operations, providing type-safe access to system information and hardware data.

## Features

- **Simple WMI Query Interface**: Easily retrieve WMI class properties
- **Type-Safe Conversions**: Automatic conversion between COM types and C++ types
- **Array Support**: Handle both scalar and array values
- **Modern C++ Interface**: Uses `std::variant`, `std::optional`, and C++20 features
- **Header-Only**: Easy integration into existing projects

## Requirements

- Windows 10/11 or Windows Server 2016+
- C++20 compatible compiler (MSVC 2022+ recommended)
- Windows SDK

## Basic Usage

### Retrieving Operating System Information
```cpp
const WindowsManagementInstrumentationClient wmi;
auto os = wmi.getProperties( L"Win32_OperatingSystem" );
const auto name = res.getProperty< std::wstring >( L"Caption" );
if(name.has_value()) {
  std::cout << name.value();
}
```
