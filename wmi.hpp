#pragma once

#include <optional>
#include <span>
#include <string>
#include <unordered_map>
#include <variant>
#include <system_error>
#include <functional>
#include <map>
#include <vector>
#include <windows.h>
#include <wbemcli.h>
#include <comdef.h>
#pragma comment(lib, "wbemuuid.lib")

#define THROW_LAST_IF(expr) if (expr) { throw utils::Exception(GetLastError()); }
#define THROW_LAST() throw utils::Exception(GetLastError())

namespace SimplerWMI {
using WmiValue = std::variant<
    // Scalar types
    bool, // CIM_BOOLEAN
    int8_t, // CIM_SINT8
    uint8_t, // CIM_UINT8
    int16_t, // CIM_SINT16
    uint16_t, // CIM_UINT16
    int32_t, // CIM_SINT32
    uint32_t, // CIM_UINT32
    int64_t, // CIM_SINT64
    uint64_t, // CIM_UINT64
    float, // CIM_REAL32
    double, // CIM_REAL64
    wchar_t, // CIM_CHAR16
    std::wstring, // CIM_STRING, CIM_DATETIME, CIM_REFERENCE

    // Array types (CIM_FLAG_ARRAY)
    std::vector< bool >,
    std::vector< int8_t >,
    std::vector< uint8_t >,
    std::vector< int16_t >,
    std::vector< uint16_t >,
    std::vector< int32_t >,
    std::vector< uint32_t >,
    std::vector< int64_t >,
    std::vector< uint64_t >,
    std::vector< float >,
    std::vector< double >,
    std::vector< wchar_t >,
    std::vector< std::wstring >
>;

namespace utils {
    class Exception : public std::exception {
    public:
        explicit Exception(const uint32_t error_code)
            : message( std::system_category().default_error_condition( error_code ).message() ) {}

        [[nodiscard]] const char *what() const noexcept override { return message.c_str(); }

    private:
        std::string message;
    };
}

class WindowsManagementInstrumentationObject;

class WindowsManagementInstrumentationClient {
public:
    WindowsManagementInstrumentationClient() {
        HRESULT hr = CoInitializeEx( nullptr, COINIT_MULTITHREADED );
        THROW_LAST_IF( FAILED(hr) );

        hr = CoCreateInstance( CLSID_WbemLocator, nullptr, CLSCTX_INPROC_SERVER,
                               IID_IWbemLocator, reinterpret_cast< LPVOID * >( &pLoc ) );
        if ( FAILED( hr ) ) {
            CoUninitialize();
            THROW_LAST();
        }

        hr = pLoc->ConnectServer( _bstr_t( L"ROOT\\CIMV2" ), nullptr, nullptr, nullptr, 0, nullptr, nullptr, &pSvc );
        if ( FAILED( hr ) ) {
            pLoc->Release();
            CoUninitialize();
            THROW_LAST();
        }

        hr = CoSetProxyBlanket(
            pSvc,
            RPC_C_AUTHN_WINNT,
            RPC_C_AUTHZ_NONE,
            nullptr,
            RPC_C_AUTHN_LEVEL_CALL,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            nullptr,
            EOAC_NONE
        );
        if ( FAILED( hr ) ) {
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();
            THROW_LAST();
        }
    }

    ~WindowsManagementInstrumentationClient() noexcept {
        if ( pSvc ) { pSvc->Release(); }
        if ( pLoc ) { pLoc->Release(); }
        CoUninitialize();
    }

    std::vector< WindowsManagementInstrumentationObject > getProperties(
        const std::wstring &object, const std::initializer_list< std::wstring > &&properties = {}) const;

private:
    static std::wstring prepQuery(
        const std::wstring &object, const std::initializer_list< std::wstring > &&properties) {
        std::wstring query = L"SELECT ";
        if ( properties.size() > 0 ) {
            for ( const auto &prop: properties ) { query += prop + L","; }
            query.pop_back();
        }
        else { query += L"*"; }
        query += L" FROM " + object;
        return query;
    }

    static WmiValue convertVariantToWmiValue(VARIANT &vtProp, CIMTYPE cimType) {
        using VariantConverter = std::function< WmiValue(VARIANT &) >;
        static const std::map< CIMTYPE, VariantConverter > cimTypeMap = {
            { CIM_BOOLEAN, [] (VARIANT &v) { return v.boolVal != VARIANT_FALSE; } },
            { CIM_SINT8, [] (VARIANT &v) { return static_cast< int8_t >( v.cVal ); } },
            { CIM_UINT8, [] (VARIANT &v) { return static_cast< uint8_t >( v.bVal ); } },
            { CIM_SINT16, [] (VARIANT &v) { return static_cast< int16_t >( v.iVal ); } },
            { CIM_UINT16, [] (VARIANT &v) { return static_cast< uint16_t >( v.uiVal ); } },
            { CIM_SINT32, [] (VARIANT &v) { return v.intVal; } },
            { CIM_UINT32, [] (VARIANT &v) { return v.uintVal; } },
            { CIM_SINT64, [] (VARIANT &v) { return v.llVal; } },
            { CIM_UINT64, [] (VARIANT &v) { return v.ullVal; } },
            { CIM_REAL32, [] (VARIANT &v) { return v.fltVal; } },
            { CIM_REAL64, [] (VARIANT &v) { return v.dblVal; } },
            { CIM_CHAR16, [] (VARIANT &v) { return static_cast< wchar_t >( v.uiVal ); } },
            { CIM_STRING, [] (VARIANT &v) { return std::wstring( v.bstrVal ? v.bstrVal : L"" ); } },
            { CIM_DATETIME, [] (VARIANT &v) { return std::wstring( v.bstrVal ? v.bstrVal : L"" ); } },
            { CIM_REFERENCE, [] (VARIANT &v) { return std::wstring( v.bstrVal ? v.bstrVal : L"" ); } }
        };

        if ( cimType & CIM_FLAG_ARRAY ) {
            CIMTYPE baseType = cimType & ~CIM_FLAG_ARRAY;
            auto it = cimTypeMap.find( baseType );
            if ( it == cimTypeMap.end() ) throw utils::Exception( E_NOTIMPL );

            if ( !( vtProp.vt & VT_ARRAY ) ) throw utils::Exception( E_INVALIDARG );

            SAFEARRAY *sa = vtProp.parray;
            LONG lBound, uBound;
            SafeArrayGetLBound( sa, 1, &lBound );
            SafeArrayGetUBound( sa, 1, &uBound );
            LONG count = uBound - lBound + 1;

            SafeArrayLock( sa );
            void *data = sa->pvData;

            auto processArray = [&]<typename T> (T *arr) {
                std::vector< std::remove_cvref_t< T > > vec;
                for ( LONG i = 0; i < count; ++i ) {
                    VARIANT elem;
                    VariantInit( &elem );
                    elem.vt = vtProp.vt & ~VT_ARRAY;
                    if constexpr ( std::is_same_v< T, BSTR > ) { elem.bstrVal = arr[ i ]; }
                    else { memcpy( &elem, &arr[ i ], sizeof( T ) ); }
                    vec.emplace_back( std::get< T >( it->second( elem ) ) );
                }
                return vec;
            };

            WmiValue result;
            switch ( baseType ) {
            case CIM_BOOLEAN: result = processArray( static_cast< VARIANT_BOOL * >( data ) );
                break;
            case CIM_SINT8: result = processArray( static_cast< int8_t * >( data ) );
                break;
            case CIM_UINT8: result = processArray( static_cast< uint8_t * >( data ) );
                break;
            case CIM_SINT16: result = processArray( static_cast< int16_t * >( data ) );
                break;
            case CIM_UINT16: result = processArray( static_cast< uint16_t * >( data ) );
                break;
            case CIM_SINT32: result = processArray( static_cast< int32_t * >( data ) );
                break;
            case CIM_UINT32: result = processArray( static_cast< uint32_t * >( data ) );
                break;
            case CIM_SINT64: result = processArray( static_cast< int64_t * >( data ) );
                break;
            case CIM_UINT64: result = processArray( static_cast< uint64_t * >( data ) );
                break;
            case CIM_REAL32: result = processArray( static_cast< float * >( data ) );
                break;
            case CIM_REAL64: result = processArray( static_cast< double * >( data ) );
                break;
            case CIM_CHAR16: result = processArray( static_cast< wchar_t * >( data ) );
                break;
            case CIM_STRING:
            case CIM_DATETIME:
            case CIM_REFERENCE: {
                auto bstrArray = static_cast< BSTR * >( data );
                std::vector< std::wstring > stringVec;
                stringVec.reserve( count );

                for ( LONG i = 0; i < count; ++i ) { stringVec.emplace_back( bstrArray[ i ] ? bstrArray[ i ] : L"" ); }
                result = stringVec;
                break;
            }
            default: SafeArrayUnlock( sa );
                throw utils::Exception( E_NOTIMPL );
            }

            SafeArrayUnlock( sa );
            return result;
        }

        if ( auto it = cimTypeMap.find( cimType ); it != cimTypeMap.end() ) { return it->second( vtProp ); }
        throw utils::Exception( E_NOTIMPL );
    }

private:
    IWbemLocator *pLoc = nullptr;
    IWbemServices *pSvc = nullptr;
};

class WindowsManagementInstrumentationObject {
public:
    template< typename T >
    std::optional< T > getProperty(const std::wstring &prop) const {
        const auto it = properties.find( prop );
        if ( it == properties.end() ) return std::nullopt;

        if ( auto *val = std::get_if< T >( &it->second ) ) { return *val; }
        return std::nullopt;
    }

    template< typename T >
    std::span< const T > getArray(const std::wstring &prop) const {
        const auto it = properties.find( prop );
        if ( it == properties.end() ) return {};

        if ( auto *vec = std::get_if< std::vector< T > >( &it->second ) ) { return *vec; }
        return {};
    }

private:
    friend class WindowsManagementInstrumentationClient;
    void addProperty(const std::wstring &name, const WmiValue &value) { properties[ name ] = value; }

private:
    std::unordered_map< std::wstring, WmiValue > properties;
};

inline std::vector< WindowsManagementInstrumentationObject > WindowsManagementInstrumentationClient::getProperties(
    const std::wstring &object, const std::initializer_list< std::wstring > &&properties) const {
    IEnumWbemClassObject *pEnumerator = nullptr;
    std::vector< WindowsManagementInstrumentationObject > results;

    try {
        const auto query = prepQuery( object, std::move( properties ) );

        HRESULT hr = pSvc->ExecQuery(
            _bstr_t( L"WQL" ),
            _bstr_t( query.c_str() ),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            nullptr,
            &pEnumerator
        );
        THROW_LAST_IF( FAILED( hr ) );

        IWbemClassObject *pclsObj = nullptr;
        ULONG uReturn = 0;

        while ( true ) {
            hr = pEnumerator->Next( WBEM_INFINITE, 1, &pclsObj, &uReturn );
            if ( hr == WBEM_S_FALSE ) break;
            THROW_LAST_IF( FAILED( hr ) );

            SAFEARRAY *pNames = nullptr;
            hr = pclsObj->GetNames( nullptr, WBEM_FLAG_ALWAYS, nullptr, &pNames );
            if ( FAILED( hr ) ) {
                pclsObj->Release();
                THROW_LAST();
            }

            LONG lLower, lUpper;
            SafeArrayGetLBound( pNames, 1, &lLower );
            SafeArrayGetUBound( pNames, 1, &lUpper );

            std::vector< std::wstring > propertyNames;
            for ( LONG i = lLower; i <= lUpper; ++i ) {
                BSTR bstrName;
                SafeArrayGetElement( pNames, &i, &bstrName );
                propertyNames.emplace_back( bstrName );
                SysFreeString( bstrName );
            }
            SafeArrayDestroy( pNames );

            WindowsManagementInstrumentationObject currentObj;
            for ( const auto &propName: propertyNames ) {
                VARIANT vtProp;
                VariantInit( &vtProp );
                CIMTYPE cimType;
                LONG flFlavor = 0;

                hr = pclsObj->Get( propName.c_str(), 0, &vtProp, &cimType, &flFlavor );
                if ( FAILED( hr ) ) {
                    VariantClear( &vtProp );
                    pclsObj->Release();
                    THROW_LAST();
                }

                currentObj.addProperty( propName, convertVariantToWmiValue( vtProp, cimType ) );
                VariantClear( &vtProp );
            }
            results.push_back( currentObj );

            pclsObj->Release();
        }

        if ( FAILED( hr ) ) { THROW_LAST_IF( FAILED( hr ) ); }
    }
    catch ( const std::runtime_error &e ) {
        if ( pEnumerator ) pEnumerator->Release();
        throw std::runtime_error( e );
    }

    if ( pEnumerator ) pEnumerator->Release();
    return results;
}
}
