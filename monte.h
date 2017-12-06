// monte.h - Monte Carlo simulation in Excel.
// Copyright (c) 2017 KALX, LLC. All rights reserved. No warranty is made.

#pragma once
#include "xll/xll.h"

#define XLL_PREFIX L"MONTE"

namespace xll {
    enum MONTE_STATE { 
        MONTE_RUN,   // active simulation
        MONTE_SPIN,  // reactive simulation
        MONTE_IDLE,  // inactive simulation
        MONTE_RESET, // prepare start of simulation
        MONTE_PAUSE  // idle during a simulation 
    };
    class monte {
        static size_t count_;
        static MONTE_STATE state_, next_state_;
        static int calculate_;
    public:
        static size_t count()
        {
            return count_;
        }
        static MONTE_STATE state()
        {
            return state_;
        }
        static MONTE_STATE next_state(MONTE_STATE state)
        {
            next_state_ = state;

            return state_;
        }
        static void calculate()
        {
            // timer.start();
            Excel(calculate_);
            // timer.stop();
        }
        static void reset()
        {
            count_ = 0;
            state_ = MONTE_RESET;
            calculate();
            state_ = next_state_ = MONTE_IDLE;
        }
        static void step()
        {
            ++count_;
            state_ = MONTE_RUN;
            calculate();
            state_ = next_state_;
        }
    };
}
