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
int main() {
    SimplerWMI::WindowsManagementInstrumentationClient client;
    const auto os = client.getProperties( L"Win32_OperatingSystem" );
    if ( !os.empty() ) {
        for ( const auto &prop: os ) {
            const auto name = prop.getProperty< std::wstring >( L"Caption" );
            if ( name.has_value() ) { std::wcout << name.value(); }
        }
    }

    const auto ram_sticks = client.getProperties( L"Win32_PhysicalMemory" );
    for ( const auto &stick: ram_sticks ) {
        if ( auto model = stick.getProperty< uint32_t >( L"Speed" ) ) {
            std::wcout << L"RAM Model: " << *model << '\n';
        }
    }
}
```
