#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#define VC_EXTRALEAN
#define _USE_MATH_DEFINES

#include <ctype.h>
#include <windows.h>
#include <math.h>
#include <malloc.h>
#include <shlwapi.h>

// XLL
#include <XLCALL.H>
#include <FRAMEWRK.H>

// PROJ
#include <proj.h>
#include <geodesic.h>
#if PROJ_VERSION_MAJOR < 8
  #define ACCEPT_USE_OF_DEPRECATED_PROJ_API_H
  #include <proj_api.h>
#endif

#include "util.h"
#include "epsg.h"

#define rgWorksheetFuncsRows 12
#define rgWorksheetFuncsCols 15

__declspec(dllexport) int WINAPI xlAutoOpen(void);
__declspec(dllexport) int WINAPI xlAutoClose(void);
__declspec(dllexport) LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName);
__declspec(dllexport) int WINAPI xlAutoAdd(void);
__declspec(dllexport) int WINAPI xlAutoRemove(void);
__declspec(dllexport) void WINAPI xlAutoFree12(LPXLOPER12 pxFree);
__declspec(dllexport) LPXLOPER12 WINAPI projVersion(LPXLOPER12 x);
__declspec(dllexport) LPXLOPER12 WINAPI projTransform(const char* src, const char* dst, const double x, const double y, const WORD type);
__declspec(dllexport) LPXLOPER12 WINAPI projTransform_api6(const char* src, const char* dst, const double x, const double y, const WORD type);
__declspec(dllexport) LPXLOPER12 WINAPI projEPSG(const int code);
__declspec(dllexport) LPXLOPER12 WINAPI projGeodInv(const char* src, const double x1, const double y1, const double x2, const double y2, const WORD type);
__declspec(dllexport) LPXLOPER12 WINAPI projGeodDir(const char* src, const double x1, const double y1, const double az1, const double dist, const WORD type);
__declspec(dllexport) LPXLOPER12 WINAPI projExec(const char* src, const double x, const double y, const double z, const double t, const WORD type);
__declspec(dllexport) LPXLOPER12 WINAPI projDeg2DMS(const double deg, const char *pos, const char *neg, const char *dchar);
__declspec(dllexport) LPXLOPER12 WINAPI projDMS2Deg(const char *dms);
__declspec(dllexport) LPXLOPER12 WINAPI projGetCRSList(char *AuthFilter, char *NameFilter, int fCol, int fRow);
__declspec(dllexport) LPXLOPER12 WINAPI projGetCRSListSize(char *AuthFilter, char *NameFilter);
__declspec(dllexport) LPXLOPER12 WINAPI projTransformInfo(char *src, char *dest, double x, double y);
