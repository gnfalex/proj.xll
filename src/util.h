#pragma once

#include <windows.h>

int lpwstricmp(LPWSTR s, LPWSTR t);
wchar_t *new_xl12string(const char *text);
char *xl12string2multibyte(const wchar_t *text, UINT cp);
int cutFileNameFromPath(char *fpath);
