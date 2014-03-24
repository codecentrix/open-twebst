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

// This file contains help messages used by the IDL file.
#pragma once

// ICore interface help messages.
#define CORE_START_BROWSER_HELP             "Starts a new instance of Internet Explorer and returns a Browser object connected to it."
#define CORE_FIND_BROWSER_HELP              "Finds an existing instance of Internet Explorer and returns a Browser object connected to it."
#define CORE_SEARCH_TIMEOUT_HELP            "The amount of time search methods try to find a browser, frame or html element."
#define CORE_LOAD_TIMEOUT_HELP              "The amount of time search methods wait the web page to load."
#define CORE_LAST_ERROR_HELP                "The last error code for any method/property call in the library."
#define CORE_LOAD_TIMEOUT_IS_ERR_HELP       "Specifies if search methods throw script exceptions when a load timeout occurs."
#define CORE_OK_ERROR_HELP                  "Error code for: \"No error occured.\""
#define CORE_FAIL_ERROR_HELP                "Error code for: \"A method / property call failed.\""
#define CORE_INVALID_ARG_ERROR_HELP         "Error code for: \"An invalid parameter was passed to a method.\""
#define CORE_LOAD_TIMEOUT_ERROR_HELP        "Error code that specifies that load timeout expired while waiting for the web page to complete."
#define CORE_INDEX_OUT_OF_BOUND_ERR_HELP    "Error code: \"An attempt was made to access an element of a collection with an index that exceeds the size of the collection.\""
#define CORE_BRWS_CONNECT_LOST_ERR_HELP     "Error code: \"A script object was disconnected from its Internet Explorer instance.\""
#define CORE_CREATE_FRAME_HELP              "Creates a Frame object from a html window (IHTMLWindow2)."
#define CORE_CREATE_ELEMENT_HELP            "Creates a Element object from a html element (IHTMLElement)."
#define CORE_CREATE_BROWSER_HELP            "Creates a Browser object from an Internet Explorer browser object (IWebBrowser)"
#define CORE_USE_IE_EVENTS_HELP             "Specifies if Internet Explorer events or hardware events will be used for input events."
#define CORE_INVALID_OPERATION_ERR_HELP     "Error code that specifies that the operation is not applicable to html element"
#define CORE_NOT_FOUND_ERR_HELP             "Error code that specifies that an operation failed because some elements were not found."
#define CORE_RESET_HELP                     "Resets all the core properties to their default values."
#define CORE_IE_VERSION_HELP                "Internet Explorer major version."
#define CORE_PROD_VERSION_HELP              "Twebst Library product version."
#define CORE_PROD_NAME_HELP                 "Twebst Library product name."
#define CORE_FOREGROUND_BRWS_HELP           "Retrieves the browser in foreground."
#define CORE_GET_CLIPBOARD_TEXT             "Retrieves the text from the clipboard (if any)."
#define CORE_SET_CLIPBOARD_TEXT             "Places text on the clipboard."
#define CORE_ASYNC_HTML_EVENTS_HELP         "Specifies if Click and InputText methods are executed asynchronouslye (non-blocking)."
#define CORE_CLOSE_BROWSER_POPUPS           "Auto-close browser popups (security warnings, script errors and auto complete)."
#define CORE_ELEMENT_FROM_POINT_HELP        "Returns the HTML element for the specified x and y screen coordinates."
#define CORE_ATTACH_TO_WND_HELP             "Creates a new Browser object starting from a 'Internet Explorer_Server', 'IEFrame' or 'TabWndClass' window."


// IBrowser interface help messages.
#define BROWSER_NATIVE_BROWSER_HELP       "Gets the IWebBrowser object for the connected Internet Explorer instance."
#define BROWSER_TITLE_HELP                "Gets the title of the browser."
#define BROWSER_URL_HELP                  "Gets the url of the browser."
#define BROWSER_APP_HELP                  "Gets the name of the EXE file of the browser."
#define BROWSER_IS_LOADING_HELP           "The property is true if the browser is still loading the html page, false otherwise."
#define BROWSER_WAIT_TO_LOAD_HELP         "Waits the browser to load and verify it matches a list of search conditions."
#define BROWSER_NAVIGATE_HELP             "Navigates to a given url."
#define BROWSER_CLOSE_HELP                "Closes the browser connected to the Browser script object."
#define BROWSER_CORE_HELP                 "Gets a reference to the parent Core object."
#define BROWSER_FIND_ELEMENT_HELP         "Finds an html element in the browser that matches a search condition."
#define BROWSER_TOP_FRAME_HELP            "Gets the top frame in the browser."
#define BROWSER_FIND_FRAME_HELP           "Finds a html frame inside the browser."
#define BROWSER_FIND_ALL_ELEM_HELP        "Finds all html elements in the browser that match a search condition."
#define BROWSER_NAVIGATION_ERROR          "Gets the http navigation error code."
#define BROWSER_SAVE_SNAPSHOT             "Saves a snapshot of the full document in the browser."
#define BROWSER_FIND_MODELESS_HTML_DIALOG "Finds a HTML dialog box and returns an \"IFrame\" object attached to it."
#define BROWSER_FIND_MODAL_HTML_DIALOG    "Retrieves an \"IFrame\" object attached to modal HTML dialog (if any)."
#define BROWSER_FIND_ELEMENT_BY_XPATH     "Finds an html element in the browser matching an XPath string."
#define BROWSER_CLOSE_POPUP               "Closes popup windows generated by page script with window.alert or window.confirm"
#define BROWSER_CLOSE_PROMPT              "Closes popup windows generated by page script with window.prompt"
#define BROWSER_GET_ATTR_HELP             "Get browser specific attribute by name"
#define BROWSER_SET_ATTR_HELP             "Set browser attribute to new value"


// IElement interface help messages.
#define ELEMENT_NATIVE_ELEMENT_HELP     "Gets the corresponding IHTMLElement object in the connected Internet Explorer instance."
#define ELEMENT_FIND_ELEMENT_HELP       "Finds an html element in the current html element and its sub-elements."
#define ELEMENT_FIND_CHILD_ELEMENT_HELP "Finds an html element in the current html element (doesn't search in sub-elements)."
#define ELEMENT_FIND_ALL_ELEMENTS_HELP  "Finds all html elements that match a search condition, inside a parent element."
#define ELEMENT_FIND_CHILDREN_ELEM_HELP "Finds html elements that are direct children of the given parent element."
#define ELEMENT_CORE_HELP               "Gets a reference to the parent Core object."
#define ELEMENT_CLICK_HELP              "Clicks the HTML element using hardware or IE events (depending on Core.useIEinputEvents property)."
#define ELEMENT_PARENT_ELEMENT_HELP     "Gets the parent element in the object hierarchy."
#define ELEMENT_NEXT_SIBLING_HELP       "Retrieves the next child of the parent for the element."
#define ELEMENT_PREVIOUS_SIBLING_HELP   "Retrieves the previous child of the parent for the element."
#define ELEMENT_PARENT_FRAME_HELP       "Gets the parent frame object for the current element."
#define ELEMENT_PARENT_BROWSER_HELP     "Gets the parent browser object for the current element."
#define ELEMENT_INPUT_TEXT_HELP         "Inputs text in html edit boxes using hardware or IE events (depending on Core.useIEevents property)."
#define ELEMENT_SELECT_HELP             "Selects options in a <select> html element (combo, list-box or multiple selectin list-box)."
#define ELEMENT_ADD_SELECTION_HELP      "Adds optoins to the current selection in a html multiple selection list-box."
#define ELEMENT_SELECT_RANGE_HELP       "Selects a range of options in a html multiple selection list-box."
#define ELEMENT_ADD_SEL_RAGE_HELP       "Adds a range of options to the current selection in a html multiple selection list-box."
#define ELEMENT_CLEAR_SELECTION_HELP    "Clears the current selection in a <select> element."
#define ELEMENT_UI_NAME_HELP            "Gets the text name of the html element."
#define ELEMENT_HIGHLIGHT_HELP          "Highlights the html element."
#define ELEMENT_FIND_PARENT_HELP        "Finds a parent html element that matches a search condition in the current document."
#define ELEMENT_TAGNAME_HELP            "Gets the tag name of the element."
#define ELEMENT_GET_ATTRIBUTE_HELP      "Retrieves the value of the specified attribute."
#define ELEMENT_SET_ATTRIBUTE_HELP      "Sets the value of the specified attribute."
#define ELEMENT_REMOVE_ATTRIBUTE_HELP   "Removes the given attribute from the element."
#define ELEMENT_IS_CHECKED_HELP         "The property is true if the radio/checkbox element is checked."
#define ELEMENT_SELECTED_OPTION_HELP    "Get the selected option in a <select> html element."
#define ELEMENT_ALL_SELECTED_OPTNS_HELP "Get the list of all selected options in a <select> html element."
#define ELEMENT_RIGHT_CLICK_HELP        "Right clicks the HTML element using hardware or IE events (depending on Core.useIEinputEvents property)."
#define ELEMENT_CHECK_HELP              "Check the radio/checkbox HTML element."
#define ELEMENT_UNCHECK_HELP            "Uncheck the radio/checkbox HTML element."


// IFrame interface help messages.
#define FRAME_NATIVE_FRAME_HELP           "Gets the corresponding IHTMLWindow2 object in the connected Internet Explorer instance."
#define FRAME_FIND_ELEMENT_HELP           "Finds an html element inside the frame and its sub-frames."
#define FRAME_FIND_CHILD_ELEMENT_HELP     "Finds an html element inside the frame (doesn't search inside sub-frames)."
#define FRAME_FIND_FRAME_HELP             "Finds a frame element inside the frame."
#define FRAME_FIND_CHILD_FRAME_HELP       "Finds a direct child frame element inside the frame."
#define FRAME_DOCUMENT_PROPERTY_HELP      "Gets the document object of the frame."
#define FRAME_FIND_ALL_ELEMENTS_HELP      "Finds all html elements inside the frame and its sub-frames."
#define FRAME_FIND_CHILDREN_ELEMENTS_HELP "Finds html elements that are direct children of the frame."
#define FRAME_CORE_HELP                   "Gets a reference to the parent Core object."
#define FRAME_PARENT_BROWSER_HELP         "Gets the parent browser object for the current frame."
#define FRAME_PARENT_FRAME_HELP           "Gets the parent frame object of the current frame."
#define FRAME_ELEMENT_HELP                "Retrieves the frame or iframe element that is hosting the window frame."
#define FRAME_TITLE_HELP                  "Gets the title of the frame."
#define FRAME_URL_HELP                    "Gets the url of the frame."


// IElementList interface help messages.
#define ELEMENTLIST_ITEM_HELP  "Gets a reference to an Element object in the ElementList collection."
#define ELEMENTLIST_COUNT_HELP "Gets the number of Element objects in the ElementList collection."
#define ELEMENTLIST_CORE_HELP  "Gets a reference to the parent Core object."
