// Constants used for global settings in registry.
#pragma once

namespace Settings
{
	const TCHAR REG_SETTINGS_KEY_NAME[]                    = _T("Software\\CodeCentrix\\OpenTwebst");
	const TCHAR HELP_FILE_RELATIVE_PATH[]                  = _T("Help\\OpenTwebst.chm");

	// Registry value names.
	const TCHAR REG_SETTINGS_INSTALLATION_DIRECTORY[]      = _T("InstallDir");
	const TCHAR REG_SETTINGS_LOAD_TIMEOUT_VALUE[]          = _T("loadTimeout");
	const TCHAR REG_SETTINGS_SEARCH_TIMEOUT_VALUE[]        = _T("searchTimeout");
	const TCHAR REG_SETTINGS_LOAD_TIMEOUT_IS_ERROR_VALUE[] = _T("loadTimeoutIsError");
	const TCHAR REG_SETTINGS_FLAG_VALUE[]                  = _T("settingsFlag");
	const TCHAR REG_SETTINGS_USE_IE_EVENTS[]               = _T("useHardwareEvents");
	const TCHAR REG_SETTINGS_AUTO_CLOSE_POPUPS_FILE[]      = _T("closePopupsFile");

	// Settings flags.
	const ULONG USE_DEFAULT_LOAD_TIMEOUT_FLAG          = 0x01;
	const ULONG USE_DEFAULT_SEARCH_TIMEOUT_FLAG        = 0x02;
	const ULONG USE_DEFAULT_LOAD_TIMEOUT_IS_ERROR_FLAG = 0x04;
	const ULONG USE_DEFAULT_USE_REG_EXP_FLAG           = 0x08;
	const ULONG USE_DEFAULT_IE_EVENTS_FLAG             = 0x10;

	// Default settings values.
	const ULONG        DEFAULT_SEARCH_TIMEOUT        = 10000;  // in TIME_SCALE (miliseconds).
	const ULONG        DEFAULT_LOAD_TIMEOUT          = 60000;  // in TIME_SCALE (miliseconds).
	const VARIANT_BOOL DEFAULT_LOAD_TIMEOUT_IS_ERR   = VARIANT_TRUE;
	const VARIANT_BOOL DEFAULT_USE_HARDWARE_EVENTS   = VARIANT_FALSE;
}
