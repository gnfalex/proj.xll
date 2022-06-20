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

// Create null-terminated Wide Bytes string from counted Unicode wchar string
wchar_t *xl12string2wbs (const wchar_t *text)
{
    wchar_t  *buff;
    size_t  len_buff;

    len_buff = text[0];
    buff = (wchar_t *) malloc ( (len_buff + 1) * sizeof (wchar_t) );
    memcpy (buff, (text + 1), (len_buff) * sizeof(wchar_t));
    buff[len_buff] = 0;
    return buff;
}

char *wbs2mbs (const wchar_t *text, UINT cp)
{
    size_t  len_inbuff, len_outbuff;
    char *outbuff;
    len_inbuff = wcslen(text);
    len_outbuff = WideCharToMultiByte(cp, 0, text, -1, NULL, 0, NULL, NULL);
    outbuff = malloc(len_outbuff);
    WideCharToMultiByte(cp, 0, text, -1, outbuff, len_outbuff, NULL, NULL);
    return outbuff;
}

int cutFileNameFromPathA(char *fpath){
  int i;
  for (i = strlen(fpath) - 1; i > 1; i--)
    if (fpath[i] == '\\') {fpath[i+1] = 0; break;}
  return i;
}

int cutFileNameFromPathW(wchar_t *fpath){
  int i;
  for (i = wcslen(fpath) - 1; i > 1; i--)
    if (fpath[i] == L'\\') {fpath[i+1] = 0; break;}
  return i;
}


int setXLLFolderAsProjDB(PJ_CONTEXT *ctx){
    wchar_t *searchDirW;
    XLOPER12 xXLL;
    Excel12f(xlGetName, &xXLL, 0);
    searchDirW = xl12string2wbs(xXLL.val.str);
    Excel12f(xlFree, 0, 1,  (LPXLOPER12) &xXLL);
    cutFileNameFromPathW(searchDirW);
    return setFolderAsProjDBW(searchDirW, ctx);
}

int setFolderAsProjDBW(wchar_t *searchDirW, PJ_CONTEXT *ctx){
    wchar_t **outFilesListW = (wchar_t **)malloc(MAX_PATHW);
    wchar_t *searchMaskW = (wchar_t *)malloc(MAX_PATHW);
    struct _wfinddata_t fdata;
    intptr_t fsearch;
    int cursorPos=0;
    char *searchDirC;
    char **outFilesListC;
    char *projDBC;
    int errCode;

    searchDirC = wbs2mbs (searchDirW,PROJ_CODEPAGE);
    if (strstr(searchDirC,proj_context_get_database_path(ctx))!=NULL)
      return 0;

    wcscpy_s(searchMaskW,MAX_PATHW,searchDirW);
    wcscat_s(searchMaskW,MAX_PATHW,L"aux*.db");

    if (-1 != (fsearch = _wfindfirst(searchMaskW,&fdata))) {
      do {
        outFilesListW[cursorPos] = malloc(MAX_PATHW);
        wcscpy_s(outFilesListW[cursorPos],MAX_PATHW,searchDirW);
        wcscat_s(outFilesListW[cursorPos],MAX_PATHW,fdata.name);
        cursorPos++;
      } while(0 == (_wfindnext(fsearch,&fdata)));
      _findclose(fsearch);
    };
    outFilesListW[cursorPos]=0;


    outFilesListC = (char **)malloc(cursorPos+1);
    outFilesListC[cursorPos--]=0;

    for (;cursorPos>=0;cursorPos--) {
      outFilesListC[cursorPos] = wbs2mbs(outFilesListW[cursorPos],CP_UTF8);
      free (outFilesListW[cursorPos]);
    }

    proj_context_set_search_paths (PJ_DEFAULT_CTX, 1, &searchDirC);
    errCode = proj_context_errno(ctx);
    if (!errCode) {
      projDBC = malloc(MAX_PATH);
      strcpy_s(projDBC,MAX_PATH,searchDirC);
      strcat_s(projDBC,MAX_PATH,"proj.db");
      proj_context_set_database_path(PJ_DEFAULT_CTX, projDBC, outFilesListC, NULL);
      errCode = proj_context_errno(ctx);
    }

    free (outFilesListW);
    free (searchMaskW);
    return errCode;
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
  proj_errno_reset(NULL);
  return res;
}
