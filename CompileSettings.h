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
	const BYTE   MINOR_VERSION   = 4;

	const TCHAR OPEN_TWEBST_PRODUCT_NAME[] = _T("Open Twebst");
}
