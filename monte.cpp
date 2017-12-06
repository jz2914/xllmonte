// monte.cpp - Monte Carlo simulation for Excel
// Copyright (c) 2017 KALX, LLC. All rights reserved. No warranty is made.
#include "monte.h"

using namespace xll;

size_t monte::count_ = 0;
MONTE_STATE monte::state_ = MONTE_IDLE;
MONTE_STATE monte::next_state_ = MONTE_IDLE;
int monte::calculate_ = xlcCalculateNow;

// screen updates every 0.5 seconds
double update = 0.5;

struct Timer {
    double start_, stop_, elapsed_;
    bool running_ = false;
    void reset()
    {
        elapsed_ = 0;
        running_ = false;
    }
    void start()
    {
        start_ = Excel(xlfNow);
        elapsed_ = 0;
        running_ = true;
    }
    void pause() {
        stop_ = Excel(xlfNow);
        elapsed_ += stop_ - start_;
        running_ = false;
    }
    void resume()
    {
        start_ = Excel(xlfNow);
        running_ = true;
    }
    // elapsed time in seconds
    double elapsed() const
    {
        return 86400.*(elapsed_ + (running_ ? Excel(xlfNow) - start_ : 0));
    }
} timer, updater;

AddIn xai_monte_count(
    Function(XLL_DOUBLE, L"?xll_monte_count_", L"MONTE.COUNT")
    .Volatile()
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
    .Volatile()
    .FunctionHelp(L"Return elapsed time in seconds since the start of the simulation.")
    .Category(XLL_PREFIX)
);
double WINAPI xll_monte_elapsed_(void)
{
#pragma XLLEXPORT
    return timer.elapsed();
}

void updating(bool b)
{
    Excel(xlcEcho, OPER(b));
    Excel(xlcEcho, OPER(!b));
    Excel(xlcEcho, OPER(b));
}

static void update_message_bar()
{
	static OPER xFmtElapsed(L"mm:ss.0;@");
	static OPER xFmtCount(L"* #,##0");
	static OPER xL(L" ["), xR(L"] "), xE(L"/s");
    
    double count = monte::count();
    double elapsed = timer.elapsed();
	double per = elapsed == 0 ? 0 : count/elapsed;

    updating(true);
	Excel(xlcMessage
		,OPER(true)
		,Excel(xlfConcatenate
			,Excel(xlfText, OPER(count), xFmtCount)
			,xL
			,Excel(xlfText, OPER(elapsed/86400.), xFmtElapsed)
			,xR
			,Excel(xlfText, OPER(per), xFmtCount)
			,xE
		)
	);
    // Fool Excel into updating workbook.
//    const auto& wb = Excel(xlfGetWorkbook, OPER(16));
//    Excel(xlcWorkbookSelect, wb);
    updating(false);

}

AddIn xai_monte_reset(Macro(L"?xll_monte_reset", XLL_PREFIX L".RESET"));
int WINAPI xll_monte_reset()
{
#pragma XLLEXPORT
    timer.reset();
    monte::reset();
    update_message_bar();

    return TRUE;
}

AddIn xai_monte_step(Macro(L"?xll_monte_step", XLL_PREFIX L".STEP"));
int WINAPI xll_monte_step()
{
#pragma XLLEXPORT
    monte::next_state(monte::state());
    monte::step();
    update_message_bar();
    
    return TRUE;
}

AddIn xai_monte_run(Macro(L"?xll_monte_run", XLL_PREFIX L".RUN"));
int WINAPI xll_monte_run()
{
#pragma XLLEXPORT
    if (monte::state() == MONTE_IDLE) {
        timer.start();
    }
    monte::next_state(MONTE_RUN);
    monte::step(); // show signs of life

    updater.start();
    updating(false);
    while (monte::state() == MONTE_RUN) {
        if (updater.elapsed() > update) {
            updater.start();
            update_message_bar();
            if (Excel(xlAbort)) {
                monte::next_state(MONTE_PAUSE);
            }
        }
        timer.resume();
        monte::step();
        timer.pause();
    }
    updating(true);
    update_message_bar();

    return TRUE;
}

AddIn xai_monte_pause(
    Function(XLL_BOOL, L"?xll_monte_pause", XLL_PREFIX L".PAUSE")
    .Arg(XLL_BOOL, L"condition", L"Pause simulation if condition is true.")
    .Category(XLL_PREFIX)
);
BOOL WINAPI xll_monte_pause(BOOL condition)
{
#pragma XLLEXPORT
    if (condition) {
        monte::next_state(MONTE_PAUSE);
    }

    return condition;
}