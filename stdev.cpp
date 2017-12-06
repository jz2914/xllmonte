// stdev.cpp - running mean and standard deviation of a cell
// Copyright (c) 2017 KALX, LLC. All rights reserved. No warranty is made.
#include "monte.h"

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
        o = Excel(xlCoerce, Excel(xlfCaller));
        MONTE_STATE state = monte::state();

        if (state == MONTE_RESET || reset) {
            o[0] = 0;
            if (o.size() > 1) {
                o[1] = 0;
                if (o.size() > 2) {
                    o[2] = 0;
                }
            }
        }

        if (state == MONTE_RUN) {
            double count = monte::count();
            if (o.size() > 2) {
                o[2] = o[2] + 1;
                count = o[2];
            }
		    o[0] = o[0] + (x - o[0])/count;
		    if (o.size() > 1) {
                if (count == 1) {
                    o[1] = 0;
                }
                else {
			        o[1] = (count - 2)*o[1]*o[1] 
			  		    + count*(x - o[0])*(x - o[0])/(count - 1);
			        o[1] = sqrt(o[1]/(count - 1));
                }
		    }
        }
    }
    catch (const std::exception& ex) {
        XLL_ERROR(ex.what());

        o = OPER(xlerr::NA);
    }

    return &o;
}