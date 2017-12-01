// xllmonte.h
#pragma once
#include "xll/xll.h"

#define XLL_PREFIX L"MONTE"

namespace xll {
    class monte {
    public:
        enum monte_state { 
            RUN,   // active simulation
            SPIN,  // reactive simulation
            IDLE,  // inactive simulation
            RESET, // prepare start of simulation
            PAUSE  // idle during a simulation 
        };
        static size_t count_; // global simulation count
        static double start_; // Excel date for start of simulation
        static monte_state curr_, next_; // current and next state

        static size_t count()
        {
            return count_;
        }
        // time in seconds since start of simulation
        static double elapsed()
        {
            return (Excel(xlfNow) - start_)*86400;
        }
        static monte_state state()
        {
            return curr_;
        }

        static void start(void)
        {
            curr_ = RESET;
            next_ = RUN;
        }
        static void step(void)
        {
            ++count_;
            curr_ = next_;
        }
        static void stop(void)
        {
            next_ = IDLE;
        }
        static void reset(void)
        {
            count_ = 0;
            start_ = Excel(xlfNow);
            curr_ = RESET;
            next_ = IDLE;
        }
        static void pause(void)
        {
            // stop timer
        }
    };
} // xll
