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

namespace Common
{
	// Class library error codes.
	const DWORD ERR_OK                       = 0x00;
	const DWORD ERR_FAIL                     = 0x01;
	const DWORD ERR_INVALID_ARG              = 0x02;
	const DWORD ERR_LOAD_TIMEOUT             = 0x03;
	const DWORD ERR_INDEX_OUT_OF_BOUND       = 0x04;
	const DWORD ERR_BRWS_CONNECTION_LOST     = 0x05;
	const DWORD ERR_OPERATION_NOT_APPLICABLE = 0x06;
	const DWORD ERR_NOT_FOUND                = 0x07;
	const DWORD ERR_CANCELED                 = 0x08;

	// Class library error HRESULTs codes.
	const HRESULT HRES_OK                       = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, ERR_OK);
	const HRESULT HRES_FAIL                     = MAKE_HRESULT(SEVERITY_ERROR,   FACILITY_ITF, ERR_FAIL);
	const HRESULT HRES_INVALID_ARG              = E_INVALIDARG;
	const HRESULT HRES_LOAD_TIMEOUT_ERR         = MAKE_HRESULT(SEVERITY_ERROR,   FACILITY_ITF, ERR_LOAD_TIMEOUT);
	const HRESULT HRES_LOAD_TIMEOUT_WARN        = MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_ITF, ERR_LOAD_TIMEOUT);
	const HRESULT HRES_INDEX_OUT_OF_BOUND_ERR   = MAKE_HRESULT(SEVERITY_ERROR,   FACILITY_ITF, ERR_INDEX_OUT_OF_BOUND);
	const HRESULT HRES_BRWS_CONNECTION_LOST_ERR = MAKE_HRESULT(SEVERITY_ERROR,   FACILITY_ITF, ERR_BRWS_CONNECTION_LOST);
	const HRESULT HRES_OPERATION_NOT_APPLICABLE = MAKE_HRESULT(SEVERITY_ERROR,   FACILITY_ITF, ERR_OPERATION_NOT_APPLICABLE);
	const HRESULT HRES_NOT_FOUND_ERR            = MAKE_HRESULT(SEVERITY_ERROR,   FACILITY_ITF, ERR_NOT_FOUND);
	const HRESULT HRES_CANCELED_ERR             = MAKE_HRESULT(SEVERITY_ERROR,   FACILITY_ITF, ERR_CANCELED);
}
