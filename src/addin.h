__declspec(dllexport) int WINAPI xlAutoOpen(void);
__declspec(dllexport) int WINAPI xlAutoClose(void);
__declspec(dllexport) LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName);
__declspec(dllexport) int WINAPI xlAutoAdd(void);
__declspec(dllexport) int WINAPI xlAutoRemove(void);
__declspec(dllexport) LPXLOPER12 WINAPI projVersion(LPXLOPER12 x);
__declspec(dllexport) LPXLOPER12 WINAPI projTransform(const char* src, const char* dst, const double x, const double y, const WORD type);
__declspec(dllexport) LPXLOPER12 WINAPI projTransform_api6(const char* src, const char* dst, const double x, const double y, const WORD type);
__declspec(dllexport) LPXLOPER12 WINAPI projEPSG(const int code);
