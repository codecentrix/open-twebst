/*
 * This file is part of Open Twebst - web automation framework.
 * Copyright (c) 2012 Adrian Dorache
 * adrian.dorache@codecentrix.com
 *
 * Open Twebst is free software: you can redistribute it
 * and/or modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * Open Twebst is distributed in the hope that it will be
 * useful, but WITHOUT ANY WARRANTY; without even the implied warranty
 * of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with Open Twebst. If not, see <http://www.gnu.org/licenses/>.
 *
 *
 * Twebst can be used under a commercial license if such has been acquired
 * (see http://www.codecentrix.com/). The commercial license does not
 * cover derived or ported versions created by third parties under GPL.
 */

// Definitions for the registry helpers functions
#include "StdAfx.h"
#include "Registry.h"


using namespace Common;

// Local function declarations.
namespace Registry
{
	HKEY  RegOpenKey(HKEY hParentKey, const String& sSubKey, REGSAM samDesired);
	DWORD RegGetStringValueLength(HKEY hParentKey, const String& sValueName);
}


// Write a string value in the registry. hMainKey is intended to be one of the predefined keys:
// HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE etc.
// The function could throw a ErrorServices::RegistryException in case of failure.
void Registry::RegSetStringValue(HKEY hMainKey, const String& sKey, const String& sValueName, const String& sValue)
{
	// Open the key hMainKey\sKey"
	_ASSERTE(hMainKey != NULL);
	HKEY hKey = Registry::RegOpenKey(hMainKey, sKey, KEY_SET_VALUE);

	// Write the value in the registry.
	LONG lResult = ::RegSetValueEx(hKey, sValueName.c_str(), 0, REG_SZ,
	                               (const BYTE*)(sValue.c_str()),
								   (DWORD)((sValue.length() + 1) * sizeof(TCHAR)));
	if (lResult != ERROR_SUCCESS)
	{
		lResult = ::RegCloseKey(hKey);
		_ASSERTE(ERROR_SUCCESS == lResult);
		throw CreateRegistryException(lResult, _T("RegSetValueEx failed in Registry::RegSetStringValue"));
	}

	// Close the registry key.
	lResult = ::RegCloseKey(hKey);
	_ASSERTE(ERROR_SUCCESS == lResult);
}


// Create a new key in the registry. hMainKey is intended to be one of the predefined keys:
// HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE etc.
// The function could throw a ErrorServices::RegistryException in case of failure.
void Registry::RegCreateKey(HKEY hMainKey, const String& sKey, const String& sNewKeyName)
{
	// Open the key hMainKey\sKey"
	_ASSERTE(hMainKey != NULL);
	HKEY hKey = Registry::RegOpenKey(hMainKey, sKey, KEY_CREATE_SUB_KEY);

	// Create the new subkey.
	HKEY hNewKey = NULL;
	LONG lResult = ::RegCreateKey(hKey, sNewKeyName.c_str(), &hNewKey);
	if ((lResult != ERROR_SUCCESS) || (NULL == hNewKey))
	{
		lResult = ::RegCloseKey(hKey);
		_ASSERTE(ERROR_SUCCESS == lResult);
		throw CreateRegistryException(lResult, _T("RegCreateKey failed in Registry::RegCreateKey"));
	}

	// Close the registry key.
	lResult = ::RegCloseKey(hKey);
	_ASSERTE(ERROR_SUCCESS == lResult);

	// Close the newly created registry key.
	lResult = ::RegCloseKey(hNewKey);
	_ASSERTE(ERROR_SUCCESS == lResult);
}


// Get a string value from the registry. hMainKey is intended to be one of the predefined keys:
// HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE etc.
// The function could throw a ErrorServices::RegistryException in case of failure.
String Registry::RegGetStringValue(HKEY hMainKey, const String& sKey, const String& sValueName)
{
	// Open the key hMainKey\sKey"
	_ASSERTE(hMainKey != NULL);
	HKEY hKey = Registry::RegOpenKey(hMainKey, sKey, KEY_READ);

	// Get the lenght in bytes of the registry value and allocate a buffer for the string.
	DWORD  dwDataLen = RegGetStringValueLength(hKey, sValueName);
	TCHAR* szBuffer  = new TCHAR[dwDataLen / sizeof(TCHAR)];

	// Query the sValueName string value under the hKey.
	DWORD   dwValueType;
	LRESULT lResult = ::RegQueryValueEx(hKey, sValueName.c_str(), NULL, &dwValueType, (LPBYTE)szBuffer, &dwDataLen);

	if((lResult != ERROR_SUCCESS) || (dwValueType != REG_SZ))
	{
		lResult = ::RegCloseKey(hKey);
		_ASSERTE(ERROR_SUCCESS == lResult);
		delete[] szBuffer;
		throw CreateRegistryException(lResult, _T("RegQueryValueEx failed in Registry::RegGetStringValue"));
	}

	lResult = ::RegCloseKey(hKey);
	_ASSERTE(ERROR_SUCCESS == lResult);
	String sValue = szBuffer;
	delete[] szBuffer;
	return sValue;
}


// Get a DWORD value from the registry. hMainKey is intended to be one of the predefined keys:
// HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE etc.
// The function could throw a ErrorServices::RegistryException in case of failure.
DWORD Registry::RegGetDWORDValue(HKEY hMainKey, const String& sKey, const String& sValueName)
{
	// Open the key hMainKey\sKey"
	_ASSERTE(hMainKey != NULL);
	HKEY hKey = Registry::RegOpenKey(hMainKey, sKey, KEY_READ);

	// Query the sValueName string value under the hKey.
	DWORD   dwValueType;
	DWORD   dwValue;
	DWORD	dwValLen = sizeof(DWORD);
	LRESULT lResult  = ::RegQueryValueEx(hKey, sValueName.c_str(), NULL, &dwValueType, (LPBYTE)(&dwValue), &dwValLen);

	if((lResult != ERROR_SUCCESS) || (dwValueType != REG_DWORD))
	{
		lResult = ::RegCloseKey(hKey);
		_ASSERTE(ERROR_SUCCESS == lResult);
		throw CreateRegistryException(lResult, _T("RegQueryValueEx failed in Registry::RegGetDWORDValue"));
	}

	lResult = ::RegCloseKey(hKey);
	_ASSERTE(ERROR_SUCCESS == lResult);
	return dwValue;
}


// Opens a registry subkey under a parent key and returns a HKEY.
// sSubKey is the name of the subkey. The functions could throw a
// ErrorServices::RegistryException in the case of a failure.
HKEY Registry::RegOpenKey(HKEY hParentKey, const String& sSubKey, REGSAM samDesired)
{
	HKEY hSubKey = NULL;
	LONG lResult = ::RegOpenKeyEx(hParentKey, sSubKey.c_str(), 0, samDesired, &hSubKey);
	if(lResult != ERROR_SUCCESS)
	{
		throw CreateRegistryException(lResult, _T("RegOpenKeyEx failed in Registry::RegOpenKey"));
	}

	return hSubKey;
}


// Get the length in bytes of the sValueName registry value.
DWORD Registry::RegGetStringValueLength(HKEY hKey, const String& sValueName)
{
	DWORD	dwValueLen;
	LRESULT lResult = ::RegQueryValueEx(hKey, sValueName.c_str(), NULL, NULL, NULL, &dwValueLen);
	if (lResult != ERROR_SUCCESS)
	{
		throw CreateRegistryException(lResult, _T("RegQueryValueEx failed in Registry::RegGetStringValueLength"));
	}

	return dwValueLen;
}
