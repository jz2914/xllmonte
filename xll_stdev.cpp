// xllmonte_stdev.cpp - running mean and standard deviation of a cell
#include "xllmonte.h"

using namespace xll;

AddIn xai_monte_stdev(
    Function(XLL_LPOPER, L"?xll_monte_stdev", XLL_PREFIX L".STDEV")
    .Arg(XLL_DOUBLE, L"x", L"is the number whos mean and stdev is to be computed.")
    .Arg(XLL_BOOL, L"reset?", L"is an optional boolean to reset the calculation.")
    .Category(XLL_PREFIX)
    .FunctionHelp(L"Calculate the stdev...")
    .Uncalced()
);
LPOPER WINAPI xll_monte_stdev(double x, BOOL reset)
{
#pragma XLLEXPORT
    static OPER o;

    try {
        double count = xll::monte::count();
        o = Excel(xlCoerce, Excel(xlfCaller));
		
		LPXLOPER12& a = o.val.array.lparray;

		if (reset) {
			a[0] = OPER(0);
			if (o.size() > 0)
				a[1] = OPER(1);
		}
		if (o.size() == 1) {
			o.val.num = x + (x - o.val.num)/count;
		}
		else if (o.size() == 2) {
			a[0].val.num += (x - a[0].val.num)/count; 
			if (count == 1)
				a[1].val.num = 0;
			else {
				a[1].val.num = (count - 2)*a[1].val.num*a[1].val.num 
					+ count*(x - a[0].val.num)*(x - a[0].val.num)/(count - 1);
				a[1].val.num = sqrt(a[1].val.num/(count - 1));
			}
		}
		else if (o.size() == 3) {
			a[2].val.num += 1;
			a[0].val.num += (x - a[0].val.num)/a[2].val.num; 
			if (a[2].val.num == 1)
				a[1].val.num = 0;
			else {
				a[1].val.num = (count - 2)*a[1].val.num*a[1].val.num 
					+ (x - a[0].val.num)*(x - a[0].val.num)/(1 - 1/a[2].val.num);
				a[1].val.num = sqrt(a[1].val.num/(a[2].val.num - 1));
			}
		}
		else {
			if (a[3].xltype == xltypeNum) {
				a[3].xltype = xltypeBool;
				a[3].val.xbool = a[2].val.num != 0;
			}
			if (a[3].val.xbool) {
				// reset
				a[0].val.num = x;
				a[1].val.num = 0;
				a[2].val.num = 1;
				a[3].val.xbool = reset;
			}
			else {
				a[2].val.num += 1;
				a[0].val.num += (x - a[0].val.num)/a[2].val.num; 
				if (a[2].val.num == 1)
					a[1].val.num = 0;
				else {
					a[1].val.num = (a[2].val.num - 2)*a[1].val.num*a[1].val.num 
						+ (x - a[0].val.num)*(x - a[0].val.num)/(1 - 1/a[2].val.num);
					a[1].val.num = sqrt(a[1].val.num/(a[2].val.num - 1));
				}
				a[3].val.xbool = reset;
			}
		}
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        o = OPER(xlerr::Value);
    }

    return &o;
}