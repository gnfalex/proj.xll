#pragma once

#include <windows.h>
#include <XLCALL.H>
#include <FRAMEWRK.H>
#include <proj.h>
#include <stdio.h>
#include <wchar.h>
#include <io.h>
#if PROJ_VERSION_MAJOR <7
  #define PROJ_CODEPAGE CP_ACP
#else
  #define PROJ_CODEPAGE CP_UTF8
#endif
#define MAX_PATHW MAX_PATH * sizeof(wchar_t)

int lpwstricmp(LPWSTR s, LPWSTR t);
wchar_t *new_xl12string(const char *text);
wchar_t *xl12string2wbs (const wchar_t *text);
int cutFileNameFromPathA(char *fpath);
int cutFileNameFromPathW(wchar_t *fpath);
char *wbs2mbs (const wchar_t *text, UINT cp);
int setXLLFolderAsProjDB(PJ_CONTEXT *ctx);
int setFolderAsProjDBW(wchar_t *searchDirW, PJ_CONTEXT *ctx);
LPXLOPER12 setError(LPXLOPER12 res, PJ_CONTEXT *ctx, int errtype, char *txt);
