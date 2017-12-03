// xllmonte_mean.cpp - running mean of a cell
#include "monte.h"

using namespace xll;

AddIn xai_monte_mean(
    Function(XLL_LPOPER, L"?xll_monte_mean", XLL_PREFIX L".MEAN")
    .Arg(XLL_DOUBLE, L"x", L"is the number whos mean is to be computed.")
    .Arg(XLL_BOOL, L"reset?", L"is an optional boolean to reset the calculation.")
    .Category(XLL_PREFIX)
    .FunctionHelp(L"Calculate the mean...")
    .Uncalced()
);
LPOPER WINAPI xll_monte_mean(double x, BOOL reset)
{
#pragma XLLEXPORT
    static OPER o;

    o = Excel(xlCoerce, Excel(xlfCaller));
    MONTE_STATE state = monte::state();

    if (state == MONTE_RESET || reset) {
        o[0] = 0;
        if (o.size() > 1) {
            o[1] = 0;
        }
    }

    if (state == MONTE_RUN) {
        double count = monte::count();
        if (o.size() > 1) {
            o[1] = o[1] + 1;
            count = o[1];
        }
        o[0] = o[0] + (x - o[0])/count;
    }

    return &o;
}