#pragma once

#include <windows.h>
#include <XLCALL.H>
#include <FRAMEWRK.H>
#include <proj.h>
#include <stdio.h>
#if PROJ_VERSION_MAJOR <7
  #define PROJ_CODEPAGE CP_ACP
#else
  #define PROJ_CODEPAGE CP_UTF8
#endif

int lpwstricmp(LPWSTR s, LPWSTR t);
wchar_t *new_xl12string(const char *text);
char *xl12string2multibyte(const wchar_t *text, UINT cp);
int cutFileNameFromPath(char *fpath);
void setXLLFolderAsProjDB();
LPXLOPER12 setError(LPXLOPER12 res, PJ_CONTEXT *ctx, int errtype, char *txt);
