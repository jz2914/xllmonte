// xllmonte_mean.cpp - running mean of a cell
#include "xllmonte.h"

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

    double count = xll::monte::count();
    o = Excel(xlCoerce, Excel(xlfCaller));
    if (count) {
        if (count == 1 || reset) {
            o[0] = x;
            if (o.size() > 1)
                o[1] = 1;
        }
        else {
            if (o.size() > 1) {
                o[1] = o[1] + 1;
                o[0] = o[0] + (x - o[0])/o[1];
            }
            else {
                o[0] = o[0] + (x - o[0])/count;
            }
        }
    }

    return &o;
}