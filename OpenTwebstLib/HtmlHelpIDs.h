#pragma once

// ICore help context IDs.
#define IDH_CORE                           1000
#define IDH_CORE_START_BROWSER             1001
#define IDH_CORE_FIND_BROWSER              1003
#define IDH_CORE_SEARCH_TIMEOUT            1004
#define IDH_CORE_LOAD_TIMEOUT              1005
#define IDH_CORE_LAST_ERROR                1006
#define IDH_CORE_LOAD_TIMEOUT_IS_ERR       1007
#define IDH_CORE_CONSTANTS                 1009
#define IDH_CORE_CREATE_FRAME              1010
#define IDH_CORE_CREATE_ELEMENT            1011
#define IDH_CORE_CREATE_BROWSER            1012
#define IDH_CORE_USE_IE_EVENTS             1013
#define IDH_CORE_CANCELATION               1015
#define IDH_CORE_RESET                     1016
#define IDH_CORE_IE_VERSION                1017
#define IDH_CORE_PROD_VERSION              1018
#define IDH_CORE_PROD_NAME                 1019
#define IDH_CORE_FOREGROUND_BRWS           1020
#define IDH_CORE_GET_CLIPBOARD_TEXT        1022
#define IDH_CORE_SET_CLIPBOARD_TEXT        1023
#define IDH_CORE_ASYNC_HTML_EVENTS         1031
#define IDH_CORE_CLOSE_BROWSER_POPUPS      1036
#define IDH_CORE_ELEMENT_FROM_POINT        1037
#define IDH_CORE_ATTACH_TO_WND             1038


// IBrowser help context IDs.
#define IDH_BROWSER                        2000
#define IDH_BROWSER_NATIVE_BROWSER         2001
#define IDH_BROWSER_TITLE                  2002
#define IDH_BROWSER_URL                    2003
#define IDH_BROWSER_ISLOADING              2004
#define IDH_BROWSER_WAIT_TO_LOAD           2005
#define IDH_BROWSER_NAVIGATE               2006
#define IDH_BROWSER_CLOSE                  2007
#define IDH_BROWSER_CORE                   2010
#define IDH_BROWSER_FIND_ELEMENT           2011
#define IDH_BORWSER_TOP_FRAME              2012
#define IDH_BROWSER_FIND_FRAME             2013
#define IDH_BROWSER_FIND_ALL_ELEMENTS      2014
#define IDH_BROWSER_NAVIGATION_ERR         2022
#define IDH_BROWSER_SAVE_SNAPSHOT          2027
#define IDH_BROWSER_FIND_MODELESS_HTML_DLG 2033
#define IDH_BROWSER_FIND_MODAL_HTML_DLG    2034
#define IDH_FIND_ALL_MODELESS_HTML_DLGS    2035
#define IDH_BROWSER_FIND_ELEM_BY_XPATH     2037
#define IDH_BROWSER_CLOSE_POPUP            2038
#define IDH_BROWSER_CLOSE_PROMPT           2039
#define IDH_BROWSER_APP                    2040
#define IDH_BROWSER_GET_ATTR               2041
#define IDH_BROWSER_SET_ATTR               2042


// IFrame help context IDs.
#define IDH_FRAME                      4000
#define IDH_FRAME_NATIVE_FRAME         4001
#define IDH_FRAME_FIND_ELEMENT         4002
#define IDH_FRAME_FIND_CHILD_ELEMENT   4003
#define IDH_FRAME_FIND_FRAME           4004
#define IDH_FRAME_FIND_CHILD_FRAME     4005
#define IDH_FRAME_DOCUMENT_PROPERTY    4006
#define IDH_FRAME_FIND_ALL_ELEMENTS    4007
#define IDH_FRAME_FIND_CHILDREN_ELEM   4008
#define IDH_FRAME_CORE                 4011
#define IDH_FRAME_PARENT_BROWSER       4012
#define IDH_FRAME_PARENT_FRAME         4013
#define IDH_FRAME_ELEMENT              4014
#define IDH_FRAME_TITLE                4016
#define IDH_FRAME_URL                  4017


// IElement help context IDs.
#define IDH_ELEMENT                        6000
#define IDH_ELEMENT_NATIVE_ELEMENT         6001
#define IDH_ELEMENT_FIND_ELEMENT           6002
#define IDH_ELEMENT_FIND_CHILD_ELEMENT     6003
#define IDH_ELEMENT_FIND_ALL_ELEMENTS      6004
#define IDH_ELEMENT_FIND_CHILDREN_ELEM     6005
#define IDH_ELEMENT_CORE                   6006
#define IDH_ELEMENT_CLICK                  6007
#define IDH_ELEMENT_PARENT_ELEMENT         6009
#define IDH_ELEMENT_NEXT_SIBLING           6010
#define IDH_ELEMENT_PREVIOUS_SIBLING       6011
#define IDH_ELEMENT_PARENT_FRAME           6012
#define IDH_ELEMENT_PARENT_BROWSER         6013
#define IDH_ELEMET_INPUT_TEXT              6014
#define IDH_ELEMENT_SELECT                 6015
#define IDH_ELEMENT_ADD_SELECTION          6016
#define IDH_ELEMENT_SELECT_RANGE           6017
#define IDH_ELEMENT_ADD_SEL_RANGE          6018
#define IDH_ELEMENT_CLEAR_SELECTION        6019
#define IDH_ELEMENT_UI_NAME                6020
#define IDH_ELEMENT_HIGHLIGHT              6021
#define IDH_ELEMENT_FIND_PARENT            6031
#define IDH_ELEMENT_TAGNAME                6032
#define IDH_ELEMENT_GET_ATTRIBUTE          6033
#define IDH_ELEMENT_SET_ATTRIBUTE          6034
#define IDH_ELEMENT_REMOVE_ATTRIBUTE       6035
#define IDH_ELEMENT_IS_CHECKED             6037
#define IDH_ELEMENT_SELECTED_OPTION        6038
#define IDH_ELEMENT_ALL_SELECTED_OPTIONS   6039
#define IDH_ELEMENT_RIGHT_CLICK            6040
#define IDH_ELEMENT_CHECK                  6042
#define IDH_ELEMENT_UNCHECK                6043


// IElementList help context IDs.
#define IDH_ELEMENT_LIST               7000
#define IDH_ELEMENT_LIST_COUNT         7001
#define IDH_ELEMENT_LIST_ITEM          7002
#define IDH_ELEMENT_LIST_CORE          7003


// Other topics.
#define IDH_SEARCH_CONDITION           8000
