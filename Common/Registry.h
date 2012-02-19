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

// Declarations of functions, classes etc that deal with Windows registry.
// All the registry functions are placed in the Registry namespace.
#pragma once
#include "Exceptions.h"
#include "Common.h"


namespace Registry
{
	using namespace Common;
	String RegGetStringValue(HKEY hMainKey, const String& sKey, const String& sValueName);
	DWORD  RegGetDWORDValue (HKEY hMainKey, const String& sKey, const String& sValueName);
	void   RegSetStringValue(HKEY hMainKey, const String& sKey, const String& sValueName, const String& sValue);
	void   RegCreateKey(HKEY hMainKey, const String& sKey, const String& sNewKeyName);
}
