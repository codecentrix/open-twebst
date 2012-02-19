// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the DBGSERV_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// DBGSERV_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef DBGSERV_EXPORTS
#define DBGSERV_API __declspec(dllexport)
#else
#define DBGSERV_API __declspec(dllimport)
#endif

// This class is exported from the Dbgserv.dll
class DBGSERV_API CDbgserv {
public:
	CDbgserv(void);
	// TODO: add your methods here.
};

extern DBGSERV_API int nDbgserv;

DBGSERV_API int fnDbgserv(void);
