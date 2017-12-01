// xllmonte.cpp - Monte Carlo simulation for Excel
#include "xllmonte.h"

using namespace xll;

double update = 0.4/86400;
size_t xll::monte::count_;
double xll::monte::start_;
xll::monte::monte_state xll::monte::curr_, xll::monte::next_;


AddIn xai_monte_count(
    Function(XLL_DOUBLE, L"?xll_monte_count_", L"MONTE.COUNT")
    .FunctionHelp(L"Return current iteration count.")
    .Category(XLL_PREFIX)
);
double WINAPI xll_monte_count_(void)
{
#pragma XLLEXPORT
    return xll::monte::count();
}
AddIn xai_monte_elapsed(
    Function(XLL_DOUBLE, L"?xll_monte_elapsed_", L"MONTE.ELAPSED")
    .FunctionHelp(L"Return elapsed time in seconds since the start of the simulation.")
    .Category(XLL_PREFIX)
);
double WINAPI xll_monte_elapsed_(void)
{
#pragma XLLEXPORT
    return xll::monte::elapsed();
}

static void update_message_bar()
{
	static OPER xFmtElapsed(L"mm:ss.0;@");
	static OPER xFmtCount(L"* #,##0");
	static OPER xL(L" ["), xR(L"] "), xE(L"/s");
    
    double count = xll::monte::count();
    double elapsed = xll::monte::elapsed();
	double per = elapsed == 0 ? 0 : count/elapsed;

	Excel(xlcMessage
		,OPER(true)
		,Excel(xlfConcatenate
			,Excel(xlfText, OPER(count), xFmtCount)
			,xL
			,Excel(xlfText, OPER(elapsed/86400), xFmtElapsed)
			,xR
			,Excel(xlfText, OPER(per), xFmtCount)
			,xE
		)
	);

    // Fool Excel into updating workbook.
//    const auto& wb = Excel(xlfGetWorkbook, OPER(16));
//    Excel(xlcWorkbookSelect, wb);
}

AddIn xai_monte_reset(Macro(L"?xll_monte_reset", XLL_PREFIX L".RESET"));
int WINAPI xll_monte_reset()
{
#pragma XLLEXPORT
    xll::monte::reset();

    update_message_bar();

    return TRUE;
}

void monte_step(void)
{
    xll::monte::step();
    Excel(xlcCalculateNow);
}
AddIn xai_monte_step(Macro(L"?xll_monte_step", XLL_PREFIX L".STEP"));
int WINAPI xll_monte_step()
{
#pragma XLLEXPORT

    monte_step();
    update_message_bar();
    
    return TRUE;
}

AddIn xai_monte_run(Macro(L"?xll_monte_run", XLL_PREFIX L".RUN"));
int WINAPI xll_monte_run()
{
#pragma XLLEXPORT
    bool run{true};
    double start = Excel(xlfNow);

    xll::monte::reset();
    xll_monte_step(); // show signs of life

    Excel(xlcEcho, OPER(false));
    while (run) {
        double now = Excel(xlfNow);
        if (now - start > update) {
            start = now;

            Excel(xlcEcho, OPER(true));
            update_message_bar();
            Excel(xlcEcho, OPER(false));

            if (Excel(xlAbort)) {
                run = false;
            }
        } else {
            monte_step();
        }
    }
    Excel(xlcEcho, OPER(true));

    return TRUE;
}
