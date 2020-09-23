
#include "StdAfx.h"

long _stdcall getAlphaValue(long _sc, long _ec, int _level)	{
    int r = GetRValue(_sc), g = GetGValue(_sc), b = GetBValue(_sc);

    return RGB(
        int(r+(double)(GetRValue(_ec)-r)/256*_level+0.5),
        int(g+(double)(GetGValue(_ec)-g)/256*_level+0.5),
        int(b+(double)(GetBValue(_ec)-b)/256*_level+0.5)
	);
}

void _stdcall alphaBlend(long _width, long _height, long _level, long _targetHDC,
						 long _sourceHDC0, long _sourceHDC1)	{

	int i;
	
	long dc[3], bmp[3];
	
	for (i=0;i<3;i++)
		dc[i] = (long)CreateCompatibleDC(NULL);
	
	for(i=0;i<3;i++)
		bmp[i] = (long)CreateCompatibleBitmap((HDC)_targetHDC, _width, _height);
	
	SelectObject((HDC)dc[0], (HGDIOBJ)bmp[0]);
	BitBlt((HDC)dc[0], 0, 0, _width, _height, (HDC)_sourceHDC0, 0, 0, SRCCOPY);

	SelectObject((HDC)dc[1], (HGDIOBJ)bmp[1]);
	BitBlt((HDC)dc[1], 0, 0, _width, _height, (HDC)_sourceHDC1, 0, 0, SRCCOPY);

	SelectObject((HDC)dc[2], (HGDIOBJ)bmp[2]);

	for (long y=0;y<=_height;y++)
        for (long x=0;x<=_width;x++)	{
            SetPixelV(
                (HDC)dc[2],
                x,
                y,
                getAlphaValue(
                    GetPixel((HDC)dc[0], x, y),
                    GetPixel((HDC)dc[1], x, y),
                    _level
				)
			);
		}

	BitBlt((HDC)_targetHDC, 0, 0, _width, _height, (HDC)dc[2], 0, 0, SRCCOPY);

	for (i=0;i<3;i++)	{
		DeleteObject((HGDIOBJ)bmp[i]);
		DeleteDC((HDC)dc[i]);
	}
}
