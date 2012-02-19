#pragma once

// Comment/uncomment the line bellow to deactivate/activate traces in release build.
//#define DISABLE_TRACE_SERVICE

#ifdef _DEBUG
	#undef DISABLE_TRACE_SERVICE
#endif

namespace ProductSettings
{
	const BYTE   BUILD_NUMBER    = 0; // Increment for each official build!
	const BYTE   MAJOR_VERSION   = 1;
	const BYTE   MINOR_VERSION   = 0;

	const TCHAR OPEN_TWEBST_PRODUCT_NAME[] = _T("Open Twebst");
}
