// monte.h
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
            Excel(calculate_);
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
            // timer.start();
            calculate();
            // timer.stop();
            state_ = next_state_;
        }
    };
}
/*
class monte {
    public:
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
*/