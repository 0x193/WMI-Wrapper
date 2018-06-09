#include "WMI.h"

WMI::WMI()
{
	this->m_bIsInizialized = false;
	this->m_bIsInitializedSecurity = false;
	this->m_pLocator = NULL;
	this->m_pService = NULL;
	this->m_pEnumerator = NULL;
	this->m_pObject = NULL;
}

WMI::~WMI()
{
	if ( IsInitialized() )
	{
		::CoUninitialize();
	}
	if ( this->m_pLocator )
		this->m_pLocator->Release();
	if ( this->m_pService )
		this->m_pService->Release();
	if ( this->m_pEnumerator )
		this->m_pEnumerator->Release();
	if ( this->m_pObject )
		this->m_pObject->Release();
	::VariantClear( &this->m_vtValue );
}

bool WMI::IsInitialized()
{
	return this->m_bIsInizialized;
}

bool WMI::IsInitializedSecurity()
{
	return this->m_bIsInitializedSecurity;
}

bool WMI::Initialize()
{
	this->m_hLastError = ::CoInitializeEx( 0, COINIT_MULTITHREADED );
	if ( FAILED( this->m_hLastError ) )
		return false;
	else
	{
		this->m_bIsInizialized = true;

		return true;
	}
}

bool WMI::InitializeSecurity()
{
	if ( !this->IsInitialized() )
		return false;

	this->m_hLastError = ::CoInitializeSecurity( NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL );
	if ( FAILED( this->m_hLastError ) )
	{
		this->~WMI();

		return false;
	}
	else
	{
		this->m_bIsInitializedSecurity = true;

		return true;
	}
}

bool WMI::CreateInstance()
{
	if ( !this->IsInitialized() || !this->IsInitializedSecurity() )
		return false;

	this->m_hLastError = CoCreateInstance( CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, ( LPVOID * )&this->m_pLocator );
	if ( FAILED( this->m_hLastError ) )
	{
		this->~WMI();

		return false;
	}
	else
		return true;
}

bool WMI::ConnectServer( const std::wstring &wstrNetworkResourse )
{
	if ( !this->IsInitialized() || !this->IsInitializedSecurity() || this->m_pLocator == NULL || wstrNetworkResourse.empty() )
		return false;

	this->m_hLastError = this->m_pLocator->ConnectServer( bstr_t( wstrNetworkResourse.c_str() ), NULL, NULL, 0, NULL, 0, 0, &this->m_pService );
	if ( FAILED( this->m_hLastError ) )
	{
		this->~WMI();

		return false;
	}
	else
		return true;
}

bool WMI::SetProxyBlanket()
{
	if ( !this->IsInitialized() || !this->IsInitializedSecurity() || this->m_pLocator == NULL || this->m_pService == NULL )
		return false;

	this->m_hLastError = ::CoSetProxyBlanket( this->m_pService, RPC_C_AUTHN_WINNT, RPC_C_AUTHN_NONE, NULL, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE );
	if ( FAILED( this->m_hLastError ) )
	{
		this->~WMI();

		return false;
	}
	else
		return true;
}

bool WMI::ExecQuery( const std::wstring &wstrQuery, const std::wstring &wstrProperty )
{
	if ( !this->IsInitialized() || !this->IsInitializedSecurity() || this->m_pLocator == NULL || this->m_pService == NULL )
		return false;

	this->m_hLastError = this->m_pService->ExecQuery( bstr_t( L"WQL" ), bstr_t( wstrQuery.c_str() ), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &this->m_pEnumerator );
	if ( FAILED( m_hLastError ) )
	{
		this->~WMI();

		return false;
	}
	else
	{
		while ( this->m_pEnumerator )
		{
			this->m_hLastError = this->m_pEnumerator->Next( WBEM_INFINITE, 1, &this->m_pObject, &this->m_uReturn );
			if ( this->m_uReturn == 0 )
			{
				break;
			}

			VARIANT vtProperty;

			this->m_hLastError = this->m_pObject->Get( wstrProperty.c_str(), 0, &vtProperty, 0, 0 );
			if ( SUCCEEDED( this->m_hLastError ) )
			{
				this->m_vtValue = vtProperty;

				::VariantClear( &vtProperty );

				this->m_pObject->Release();

				break;
			}

			::VariantClear( &vtProperty );

			this->m_pObject->Release();
		}

		return true;
	}
}

void WMI::Uninitialize()
{
	this->~WMI();
}

void WMI::GetValue( VARIANT &vtBuffer )
{
	vtBuffer = this->m_vtValue;
}