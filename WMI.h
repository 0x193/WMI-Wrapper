// Author: https://t.me/xorb1n

#ifndef WMI_H
#define WMI_H

#include <windows.h>
#include <comdef.h>
#include <wbemidl.h>

#pragma comment( lib, "wbemuuid.lib" )

#include <string>

class WMI
{
	HRESULT m_hLastError;
	bool m_bIsInizialized;
	bool m_bIsInitializedSecurity;
	IWbemLocator *m_pLocator;
	IWbemServices *m_pService;
	IEnumWbemClassObject *m_pEnumerator;
	IWbemClassObject *m_pObject;
	ULONG m_uReturn;
	VARIANT m_vtValue;
public:
	WMI();
	~WMI();
	bool IsInitialized();
	bool IsInitializedSecurity();
	bool Initialize();
	bool InitializeSecurity();
	bool CreateInstance();
	bool ConnectServer( const std::wstring &wstrNetworkResourse );
	bool SetProxyBlanket();
	bool ExecQuery( const std::wstring &wstrQuery, const std::wstring &wstrProperty );
	void GetValue( VARIANT &vtBuffer );
	void Uninitialize();
};

#endif // WMI_H