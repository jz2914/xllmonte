// xllmonte.h
#pragma once
#include "xll/xll.h"

#define XLL_PREFIX L"MONTE"

namespace xll {
    class monte {
    public:
        enum state { 
            RUN,   // active simulation
            IDLE,  // inactive simulation
            RESET, // prepare start of simulation
            PAUSE  // idle during a simulation 
        };
        static size_t count_; // global simulation count
        static double start_; // Excel date for start of simulation
        static state state, state_; // current and next state

        static size_t count()
        {
            return count_;
        }
        // time in seconds since start of simulation
        static double elapsed()
        {
            return (Excel(xlfNow) - start_)*86400;
        }
        static state state()
        {
            return state_;
        }
    };
} // xll
