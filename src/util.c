#pragma warning (disable: 4996) // _CRT_SECURE_NO_WARNINGS

#include "util.h"

// Compares a pascal string and a null-terminated C-string to see if they
// are equal. Case insensitive.
int lpwstricmp(LPWSTR s, LPWSTR t)
{
    int i;

    if (wcslen(s) != *t)
        return 1;

    for (i = 1; i <= s[0]; i++)
    {
        if (towlower(s[i - 1]) != towlower(t[i]))
            return 1;
    }

    return 0;
}

// Create counted Unicode wchar string from null-terminated ASCII input
wchar_t *new_xl12string(const char *text)
{
    size_t len = strlen(text);
    if (!text || !len)
        return NULL;
    if (len > 255)
        len = 255; // truncate
    wchar_t *p = (wchar_t *)malloc((len + 2) * sizeof(wchar_t));
    if (!p) return NULL;
    mbstowcs(p + 1, text, len);
    p[0] = (wchar_t)len; // string p[1] is NOT null terminated
    p[len + 1] = 0; // now it is
    return p;
}

// Create null-terminated MultiByte string from counted Unicode wchar string
char *xl12string2multibyte (const wchar_t *text, UINT cp)
{
    wchar_t  *inbuff;
    size_t  len_inbuff, len_outbuff;
    char *outbuff;

    len_inbuff = text[0];
    inbuff = (wchar_t *) malloc ( (len_inbuff + 2) * sizeof (wchar_t) );
    memcpy (inbuff, (text + 1), (len_inbuff) * sizeof(wchar_t));
    inbuff[len_inbuff] = 0;

    // Convert to Encoding ()
    len_outbuff = WideCharToMultiByte(cp, 0, inbuff, -1, NULL, 0, NULL, NULL);
    outbuff = malloc(len_outbuff);
    WideCharToMultiByte(cp, 0, inbuff, -1, outbuff, len_outbuff, NULL, NULL);

    free(inbuff);
    return outbuff;
}

int cutFileNameFromPath(char *fpath){
  int i;
  for (i = strlen(fpath) - 1; i > 1; i--)
    if (fpath[i] == '\\' || fpath[i] == 0)
      {fpath[i] = 0; break;}
  return i;
}

void setXLLFolderAsProjDB(){
    XLOPER12 xXLL;
    char * cDir;
    Excel12f(xlGetName, &xXLL, 0);
    cDir = xl12string2multibyte(xXLL.val.str, PROJ_CODEPAGE);
    cutFileNameFromPath(cDir);
    proj_context_set_search_paths (PJ_DEFAULT_CTX, 1, &cDir);
    Excel12f(xlFree, 0, 1,  (LPXLOPER12) &xXLL);
    free(cDir);
}

LPXLOPER12 setError(LPXLOPER12 res, PJ_CONTEXT *ctx, int errtype, char *txt) {
  if (errtype == xlerrNull) {
    char buff[1024];
    sprintf_s(buff, 1024, "#%s (%s)", proj_context_errno_string(ctx, proj_context_errno(ctx)), txt);
    res->xltype = xltypeStr;
    res->val.str = new_xl12string(buff);
  } else {
    res->xltype = xltypeErr;
    res->val.err = errtype;
  }
  return res;
}
