// when.cpp - conditional execution
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "monte.h"

using namespace xll;

#ifdef _T
#undef _T
#endif
#define _T(x) L##x

static AddIn xai_monte_when(
	Function(XLL_LPXLOPER, _T("?xll_monte_when"), _T("MONTE.WHEN"))
	.Arg(XLL_BOOL, _T("Condition"), _T("is a boolean value indicating whether to evaluate the second argument"))
	.Arg(XLL_LPXLOPER, _T("Range"), _T("is a formula or reference "))
    .Uncalced()
	.Category(XLL_PREFIX)
	.FunctionHelp(_T("When Condition is true return Range, otherwise do nothing."))
	.Documentation(
		_T("This is the same as <codeInline>MONTE.CONDITIONAL(Range, Condition)</codeInline>.")
	)
);
LPXLOPER12 WINAPI
xll_monte_when(BOOL b, LPXLOPER12 px)
{
#pragma XLLEXPORT
	static OPER x;

	if (b)
		return px;

    x = Excel(xlCoerce, Excel(xlfCaller));

	return &x;
}
/*
static xcstr X_(xav_monte_conditional)[] = {
	_T("is a formula or reference."),
	_T("is a boolean value indicating whether to evaluate the second argument. "),
};
*/
static AddIn xai_monte_conditional(
	Function(XLL_LPXLOPER, _T("?xll_monte_conditional"), _T("MONTE.CONDITIONAL"))
	.Arg(XLL_LPXLOPER, _T("Range"), _T("is a formula or reference "))
	.Arg(XLL_BOOL, _T("Condition"), _T("is a boolean value indicating whether to evaluate the second argument"))
	.Category(XLL_PREFIX)
	.FunctionHelp(_T("Return Range when Condition is true, otherwise do nothing."))
    .Uncalced()
	.Documentation(
		_T("This is the same as <codeInline>MONTE.WHEN(Condition, Range)</codeInline>.")
	)
);
LPXLOPER12 WINAPI
xll_monte_conditional(LPXLOPER12 px, BOOL b)
{
#pragma XLLEXPORT
	return xll_monte_when(b, px);
}