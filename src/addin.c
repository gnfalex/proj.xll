#define NOMINMAX
#define WIN32_LEAN_AND_MEAN
#define VC_EXTRALEAN
#define _USE_MATH_DEFINES

#include <ctype.h>
#include <windows.h>
#include <math.h>

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
#include "addin.h"

#define rgWorksheetFuncsRows 7
#define rgWorksheetFuncsCols 15

static LPWSTR rgWorksheetFuncs[rgWorksheetFuncsRows][rgWorksheetFuncsCols] =
{
    {
        L"projVersion",                         // LPXLOPER12 pxProcedure
        L"UU",                                  // LPXLOPER12 pxTypeText
        L"PROJ.VERSION",                        // LPXLOPER12 pxFunctionText
        L"",                                    // LPXLOPER12 pxArgumentText
        L"1",                                   // LPXLOPER12 pxMacroType
        L"PROJ",                                // LPXLOPER12 pxCategory
        L"",                                    // LPXLOPER12 pxShortcutText
        L"",                                    // LPXLOPER12 pxHelpTopic
        L"Returns the PROJ library version.",   // LPXLOPER12 pxFunctionHelp
        L"",                                    // LPXLOPER12 pxArgumentHelp1
        L"",                                    // LPXLOPER12 pxArgumentHelp2
        L"",                                    // LPXLOPER12 pxArgumentHelp3
        L"",                                    // LPXLOPER12 pxArgumentHelp4
        L"",                                    // LPXLOPER12 pxArgumentHelp5
        L""                                     // LPXLOPER12 pxArgumentHelp6
    },
    {
        L"projTransform",
        L"UCCBBH",
        L"PROJ.TRANSFORM",
        L"",
        L"1",
        L"PROJ",
        L"",
        L"",
        L"Transform X / Y points from source coordinate system to destination coordinate system.",
        L"Source coordinate system",
        L"Destination coordinate system",
        L"X coordinate",
        L"Y coordinate",
        L"Output flag: 1 = X, 2 = Y",
        L""
    },
    {
        L"projTransform_api6",
        L"UCCBBH",
        L"PROJ.TRANSFORMV6",
        L"",
        L"1",
        L"PROJ",
        L"",
        L"",
        L"Transform X / Y points from source coordinate system to destination coordinate system.",
        L"Source coordinate system",
        L"Destination coordinate system",
        L"X coordinate",
        L"Y coordinate",
        L"Output flag: 1 = Longitude 2 = Latitude, 3 = X, 4 = Y",
        L""
    },
    {
        L"projEPSG",
        L"UJ",
        L"EPSG",
        L"",
        L"1",
        L"PROJ",
        L"",
        L"",
        L"Returns the PROJ.4 string associated with an EPSG code.",
        L"EPSG code",
        L"",
        L"",
        L"",
        L"",
        L""
    },
    {
        L"projGeodInv",
        L"UCBBBBH",
        L"PROJ.GEOD_INV",
        L"",
        L"1",
        L"PROJ",
        L"",
        L"",
        L"Show distance, azimuth and reverse azimuth between two points on geod.",
        L"Coordinate system",
        L"X1 coordinate",
        L"Y1 coordinate",
        L"X2 coordinate",
        L"Y2 coordinate",
        L"Output flag: 1 = Distance 2 = Azimuth, 3 = Reverse azimuth"
    },
    {
        L"projGeodDir",
        L"UCBBBBH",
        L"PROJ.GEOD_DIR",
        L"",
        L"1",
        L"PROJ",
        L"",
        L"",
        L"Show coordinates of point by distance and azimuth from another point.",
        L"Coordinate system",
        L"X coordinate",
        L"Y coordinate",
        L"Azimuth",
        L"Distance",
        L"Output flag: 1 = Longitude 2 = Latitude"
    },
    {
        L"projExec",
        L"UCBBBBH",
        L"PROJ.EXEC",
        L"",
        L"1",
        L"PROJ",
        L"",
        L"",
        L"Execute PROJ4 string",
        L"PROJ4 string",
        L"X coordinate",
        L"Y coordinate",
        L"Height",
        L"Epoch",
        L"Output flag: 1= Longitude 2 = Latitude, 3 = Height, 4 = Epoch"
    }
};

/*
** Standard XLL functions:
** - xlAutoOpen
** - xlAutoClose
** - xlAutoRegister12
** - xlAutoAdd
** - xlAutoRemove
** - xlAddInManagerInfo12
**
** UDFs:
** - projTransform
** - projTransform_api6
** - projVersion
** - projEPSG
** - projGeodInv
** - projGeodFor
** - projExec
*/

// Excel calls xlAutoOpen when it loads the XLL.
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    static XLOPER12 xDLL;   /* name of this DLL */
    int i;                  /* Loop index */

    /*
    ** In the following block of code the name of the XLL is obtained by
    ** calling xlGetName. This name is used as the first argument to the
    ** REGISTER function to specify the name of the XLL. Next, the XLL loops
    ** through the rgFuncs[] table, registering each function in the table using
    ** xlfRegister. Functions must be registered before you can add a menu
    ** item.
    */

    Excel12f(xlGetName, &xDLL, 0);

    for (i=0; i < rgWorksheetFuncsRows; i++)
    {
        Excel12f(xlfRegister, 0, 1 + rgWorksheetFuncsCols,
            (LPXLOPER12)&xDLL,
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][0]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][1]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][2]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][3]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][4]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][5]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][6]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][7]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][8]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][9]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][10]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][11]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][12]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][13]),
            (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][14]));
    }

    /* Free the XLL filename */
    Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

    return 1;
}

// Excel calls xlAutoClose when it unloads the XLL.
__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    int i;

     // Delete all names added by xlAutoOpen or xlAutoRegister.
    for (i = 0; i < rgWorksheetFuncsRows; i++)
    {
        Excel12f(xlfSetName, 0, 1, TempStr12(rgWorksheetFuncs[i][2]));
    }

    return 1;
}


// Excel calls xlAutoRegister12 if a macro sheet tries to register
// a function without specifying the type_text argument.
__declspec(dllexport) LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
    static XLOPER12 xDLL, xRegId;
    int i;

    /*
    ** This block initializes xRegId to a #VALUE! error first. This is done in
    ** case a function is not found to register. Next, the code loops through the
    ** functions in rgFuncs[] and uses lpstricmp to determine if the current
    ** row in rgFuncs[] represents the function that needs to be registered.
    ** When it finds the proper row, the function is registered and the
    ** register ID is returned to Microsoft Excel. If no matching function is
    ** found, an xRegId is returned containing a #VALUE! error.
    */

    xRegId.xltype = xltypeErr;
    xRegId.val.err = xlerrValue;

    for (i = 0; i < rgWorksheetFuncsRows; i++)
    {
        if (!lpwstricmp(rgWorksheetFuncs[i][0], pxName->val.str))
        {
            Excel12f(xlGetName, &xDLL, 0);

            Excel12f(xlfRegister, 0, 4,
                (LPXLOPER12)&xDLL,
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][0]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][1]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][2]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][3]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][4]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][5]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][6]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][7]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][8]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][9]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][10]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][11]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][12]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][13]),
                (LPXLOPER12)TempStr12(rgWorksheetFuncs[i][13]));

            /* Free the XLL filename */
            Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

            return (LPXLOPER12)&xRegId;
        }
    }

    //Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
    //for UDFs declared as thread safe, use alternate memory allocation mechanisms

    return (LPXLOPER12)&xRegId;
}

// When you add an XLL to the list of active add-ins, the Add-in
// Manager calls xlAutoAdd() and then opens the XLL, which in turn
// calls xlAutoOpen.
__declspec(dllexport) int WINAPI xlAutoAdd(void)
{
    return 1;
}

// When you remove an XLL from the list of active add-ins, the
// Add-in Manager calls xlAutoRemove() and then
// UNREGISTER("SAMPLE.XLL").
__declspec(dllexport) int WINAPI xlAutoRemove(void)
{
    return 1;
}

// ----------------------------------------------------------------------------
// PROJ.VERSION
// ----------------------------------------------------------------------------

__declspec(dllexport) LPXLOPER12 WINAPI projVersion(LPXLOPER12 x)
{
    static XLOPER12 xResult;

    xResult.xltype = xltypeStr;
#if PROJ_VERSION_MAJOR < 8
    xResult.val.str = new_xl12string(pj_get_release());
#else
    xResult.val.str = new_xl12string(proj_info().release);
#endif
    return (LPXLOPER12)&xResult;
}

// ----------------------------------------------------------------------------
// PROJ.TRANSFORM
// ----------------------------------------------------------------------------

__declspec(dllexport) LPXLOPER12 WINAPI projTransform(const char* src, const char* dst, const double x, const double y, const WORD type)
{
#if PROJ_VERSION_MAJOR >= 8
    return projTransform_api6(src, dst, x, y, type);
#else
    static XLOPER12 xResult;

    setXLLFolderAsProjDB();

    projPJ proj_src, proj_dst;
    proj_src = pj_init_plus(src);
    proj_dst = pj_init_plus(dst);

    if (!proj_src || !proj_dst)
    {
        xResult.xltype = xltypeErr;
        xResult.val.err = xlerrValue;
        return (LPXLOPER12)&xResult;
    }

    double x1 = x;
    double y1 = y;

    if (pj_transform(proj_src, proj_dst, 1, 1, &x1, &y1, NULL) == 0)
    {
        if (type == 1) {
            xResult.xltype = xltypeNum;
            xResult.val.num = x1;
        }
        else if (type == 2) {
            xResult.xltype = xltypeNum;
            xResult.val.num = y1;
        }
        else {
            xResult.xltype = xltypeErr;
            xResult.val.err = xlerrValue; // Invalid argument
        }
    }
    else
    {
        xResult.xltype = xltypeErr;
        xResult.val.err = xlerrNum; // Error in pj_transform
    }

    if (proj_src != NULL)
        pj_free(proj_src);
    if (proj_dst != NULL)
        pj_free(proj_dst);

    return (LPXLOPER12)&xResult;
#endif
}

// ----------------------------------------------------------------------------
// PROJ.TRANSFORM6
// ----------------------------------------------------------------------------

__declspec(dllexport) LPXLOPER12 WINAPI projTransform_api6(const char* src, const char* dst, const double x, const double y, const WORD type)
{
    static XLOPER12 xResult;
    PJ *P;
    PJ_COORD c, c_out;

    setXLLFolderAsProjDB();

    P = proj_create_crs_to_crs(PJ_DEFAULT_CTX,
                               src,
                               dst,
                               NULL);
    if (P==0)
      return (LPXLOPER12) setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Cannot create PROJ");
/*
    PJ* P_for_GIS = proj_normalize_for_visualization(PJ_DEFAULT_CTX, P);
    if (P_for_GIS !=0) {proj_destroy(P);P = P_for_GIS;}
*/
    c = proj_coord(x, y ,0, HUGE_VAL);
    c_out = proj_trans(P, PJ_FWD, c);

    if (c_out.xyzt.x == HUGE_VAL)
      {proj_destroy(P); return (LPXLOPER12)setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Impossible result value");}

    xResult.xltype = xltypeNum;
    switch (type){
      case 1: xResult.val.num = c_out.lp.lam; break;
      case 2: xResult.val.num = c_out.lp.phi; break;
      default:
        setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Unknown output type");
    }
    proj_destroy(P);

    return (LPXLOPER12)&xResult;
}

// ----------------------------------------------------------------------------
// EPSG
// ----------------------------------------------------------------------------

__declspec(dllexport) LPXLOPER12 WINAPI projEPSG(const int code)
{
    static XLOPER12 xResult;

    wchar_t *projStr = epsgLookup(code);

    if (projStr != NULL) {
        xResult.xltype = xltypeStr;
        xResult.val.str = projStr;
    }
    else {
        xResult.xltype = xltypeErr;
        xResult.val.err = xlerrNA; // No value available
    }

    return (LPXLOPER12)&xResult;
}

// ----------------------------------------------------------------------------
// geod
// ----------------------------------------------------------------------------

__declspec(dllexport) LPXLOPER12 WINAPI projGeodInv(const char* src, const double x1, const double y1, const double x2, const double y2, const WORD type)
{
    static XLOPER12 xResult;
    PJ *P,*Ellips;
    struct geod_geodesic g;
    double a, invf, dist=0, az1=0, az2=0;

    setXLLFolderAsProjDB();

    P = proj_create(PJ_DEFAULT_CTX,src);
    if (P==0)
      return (LPXLOPER12)setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Cannot create PROJ");

    Ellips = proj_get_ellipsoid(PJ_DEFAULT_CTX,P);
    if (Ellips==0)
      {proj_destroy(P); return (LPXLOPER12) setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Cannot extract ellips from PROJ");}

    proj_ellipsoid_get_parameters(PJ_DEFAULT_CTX, Ellips, &a, 0, 0, &invf);
    proj_destroy(P); proj_destroy(Ellips);

    geod_init(&g, a, 1/invf);
    geod_inverse(&g, x1, y1, x2, y2, &dist, &az1, &az2);

    if (dist == HUGE_VAL)
      return (LPXLOPER12)setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Impossible result value");

    xResult.xltype = xltypeNum;
    switch (type){
      case 1: xResult.val.num = dist; break;
      case 2: xResult.val.num = az1 ; break;
      case 3: xResult.val.num = az2 ; break;
      default:
        setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Unknown output type");
    }
    return (LPXLOPER12)&xResult;
}


__declspec(dllexport) LPXLOPER12 WINAPI projGeodDir(const char* src, const double x1, const double y1, const double az1, const double dist, const WORD type)
{
    static XLOPER12 xResult;
    PJ *P, *Ellips;
    struct geod_geodesic g;
    double a, invf, x2, y2, az2;

    setXLLFolderAsProjDB();

    P = proj_create(PJ_DEFAULT_CTX,src);
    if (P==0)
      return (LPXLOPER12)setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Cannot create PROJ");

    Ellips = proj_get_ellipsoid(PJ_DEFAULT_CTX,P);
    if (Ellips==0) {proj_destroy(P); return (LPXLOPER12)setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Cannot extract ellips from PROJ");}

    proj_ellipsoid_get_parameters(PJ_DEFAULT_CTX, Ellips, &a, 0, 0, &invf);
    proj_destroy(P); proj_destroy(Ellips);

    geod_init(&g, a, 1/invf);
    geod_direct(&g, x1, y1, az1, dist, &x2, &y2, &az2);

    if (x2 == HUGE_VAL)
      return (LPXLOPER12)setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Impossible result value");

    xResult.xltype = xltypeNum;
    switch (type){
      case 1: xResult.val.num = x2; break;
      case 2: xResult.val.num = y2; break;
      default:
        setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Cannot create PROJ");
    }

    return (LPXLOPER12)&xResult;
}

__declspec(dllexport) LPXLOPER12 WINAPI projExec(const char* src, const double x, const double y, const double z, const double t, const WORD type)
{
    static XLOPER12 xResult;
    PJ *P;
    PJ_COORD c, c_out;

    setXLLFolderAsProjDB();

    P = proj_create(PJ_DEFAULT_CTX,src);
    if (P==0)
      return (LPXLOPER12)setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Cannot create PROJ");
/*
    PJ* P_for_GIS = proj_normalize_for_visualization(PJ_DEFAULT_CTX, P);
    if (P_for_GIS !=0)  {proj_destroy(P);P = P_for_GIS;}
*/
    c = proj_coord(x, y, z, t);
    if (proj_angular_input (P, PJ_FWD)) {
      c.lpzt.lam = proj_torad (c.lpzt.lam);
      c.lpzt.phi = proj_torad (c.lpzt.phi);
    }

    c_out = proj_trans(P, PJ_FWD, c);
    if (c_out.xyzt.x == HUGE_VAL)
      {proj_destroy(P); return (LPXLOPER12)setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Impossible result value");}
    if (proj_angular_output (P, PJ_FWD)) {
        c_out.lpzt.lam =  proj_todeg (c_out.lpzt.lam);
        c_out.lpzt.phi =  proj_todeg (c_out.lpzt.phi);
    }

    xResult.xltype = xltypeNum;
    switch (type){
      case 1: xResult.val.num = c_out.xyzt.x; break;
      case 2: xResult.val.num = c_out.xyzt.y; break;
      case 3: xResult.val.num = c_out.xyzt.z; break;
      case 4: xResult.val.num = c_out.xyzt.t; break;
      default:
        setError(&xResult, PJ_DEFAULT_CTX, xlerrNull, "Unknown output type");
    }
    proj_destroy(P);
    return (LPXLOPER12)&xResult;
}
