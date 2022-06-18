proj.xll
========

Transform coordinates between various map projections directly within Excel.

This program currently uses version 4 of the [PROJ Cartographic Projections Library](https://proj.org).

Installation
------------

Copy proj.xll to `%USERPROFILE%\AppData\Roaming\Microsoft\AddIns\`.

Go to Excel Options, Add-ins, Manage Excel Add-ins. Click Browse, and select proj.xll.

Usage
-----

To display the PROJ library version:

```
=PROJ.VERSION()
```

To transform coordinates, for example from latitude/longitude to UTM zone 17:
```
=PROJ.TRANSFORM("+proj=latlong +datum=WGS84","+proj=utm +zone=17 +datum=WGS84",RADIANS(-80),RADIANS(40),1)
=PROJ.TRANSFORM("+proj=latlong +datum=WGS84","+proj=utm +zone=17 +datum=WGS84",RADIANS(-80),RADIANS(40),2)
```

Arguments:
 
1. Source coordinate system
2. Destination coordinate system
3. X value
4. Y value
5. Output: 1=X, 2=Y

Latitude and longitude values must be in radians.

EPSG codes are also supported using the `EPSG()` helper function. For example, to transform from latitude/longitude to Web Mercator:
```
=PROJ.TRANSFORM(EPSG(4326),EPSG(3857),RADIANS(-89.83152841),RADIANS(40.91627447),1)
=PROJ.TRANSFORM(EPSG(4326),EPSG(3857),RADIANS(-89.83152841),RADIANS(40.91627447),2)
```

New functions depends from proj.db and (optionaly) aux\*.db near xll. Accept as input PROJ string, WKT or datum.



To transform coordinates, for example from latitude/longitude to UTM zone 17:
```
=PROJ.TRANSFORMv6("EPSG:4326","+proj=utm +zone=17 +datum=WGS84",RADIANS(-80),RADIANS(40),1)
```
Arguments:
1. Source coordinate system
2. Destination coordinate system
3. X value
4. Y value
5. Output: 1=X, 2=Y

To find distance (or azimuth) between points:
```
=PROJ.GEOD_INV("EPSG:4326",52,45,53,45,1)
```
Arguments:
1. Coordinate system
2. X1 coordinate
3. Y1 coordinate
4. X2 coordinate
5. Y2 coordinate
6. Output flag: 1 = Distance 2 = Azimuth, 3 = Reverse azimuth

To find coordinates of second point (with known distance and azimuth):
```
=PROJ.GEOD_DIR("EPSG:4326",52,45,45,100,1)
```
Arguments:
1. Coordinate system
2. X1 coordinate
3. Y1 coordinate
4. Azimuth
5. Distance
6. Output flag: 1 = Longitude 2 = Latitude

To apply PROJ string to coordinates
```
=PROJ.EXEC("+proj=pipeline +ellps=GRS80 +step +proj=cart +step +proj=helmert +x=10 +y=3 +z=1 +step +proj=cart +inv +step +proj=merc";55;42;100;2022;1)
```
Arguments:
1. PROJ4 string
2. X coordinate
3. Y coordinate
4. Height
5. Epoch
6. Output flag: 1= Longitude 2 = Latitude, 3 = Height, 4 = Epoch

To convert degrees to string representation of degrees, minutes and seconds.
```
=PROJ.DEG2DMS (53.133,"E","W","d")
```
Arguments:
1. Degrees
2. Negative Char (optional, N or E)
3. Positive Char (optional, S or W)
4. Degree Char (optional, � as default)

```
=PROJ.DMS2DEG("45d12'45")
=PROJ.DMS2DEG("45�12'45")
```
Arguments:
1. String

Development
-----------

Download the [PROJ source](https://proj.org/download.html), tested with v4.9.3 (2016-09-02). Extract to C:\PROJ.

Edit the PROJ.4 OPTFLAGS in nmake.opt as follows, to use the static multithreaded version of the CRT library (`/MT`):

```
!IFNDEF OPTFLAGS
!IFNDEF DEBUG
OPTFLAGS=	/nologo /Ox /Op /MT
!ELSE
OPTFLAGS=	/nologo /Zi /MTd /Fdproj.pdb
!ENDIF
!ENDIF
```

Build the PROJ.4 static library from a Visual Studio command prompt:

```
C:\PROJ>nmake /f makefile.vc
C:\PROJ>nmake /f makefile.vc install-all
```

Download and install the [Excel 2013 SDK](https://www.microsoft.com/en-us/download/details.aspx?id=35567) to the default location. As with PROJ, update the SAMPLES/FRAMEWRK makefile to use the static CRT (`/MT`):

```
!if "$(TYPE)" == "DEBUG"
CPPFLAGS        =/Od /W3 /WX /EHsc /Zi /MTd /Fd"$(FRAMEWORK_PDB)" /Fo"$(FRAMEWORKBUILDDIR)\\"
!else
CPPFLAGS        =/W3 /WX /EHsc /MT /Fo"$(FRAMEWORKBUILDDIR)\\"
!endif
```

Build the library from a Visual Studio command prompt:

```
C:\2013 Office System Developer Resources\Excel2013XLLSDK\SAMPLES\FRAMEWRK>nmake TYPE=RELEASE
```

You should now be able to build the XLL using the nmake makefile.
