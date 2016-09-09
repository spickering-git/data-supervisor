using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCollectionForRealtime
{
    static class DataCollectionConstants
    {
        public const long CQG_DATA_ERROR_CODE = 2147483648;

        public const int MINUTES_IN_DAY = 1440;

        public const double EPSILON = 0.000000001;

        public const double OPTION_ZERO_PRICE = 0.000001;

        public const double OPTION_DELTA_MULTIPLIER = 100;

        public const double OPTION_ACCEPTABLE_BID_ASK_SPREAD = 20;

        public const int STRIKE_COUNT_FOR_DEFAULT_TO_THEORETICAL = 4;

        public const double OPTION_DEFAULT_THEORETICAL_PRICE_RANGE = 0.17;


    }

    public enum OPTION_EXPRESSION_TYPES
    {
        OPTION_EXPRESSION_RISK_FREE_RATE,
        SPREAD_LEG_PRICE,
    };

    public enum OPTION_SPREAD_CONTRACT_TYPE
    {
        CALL,
        PUT,
        FUTURE,
        BLANK
    };

    public enum CQG_REFRESH_STATE
    {
        NOTHING,
        DATA_UPDATED
    }

    public enum REALTIME_PRICE_FILL_TYPE
    {
        PRICE_DEFAULT,
        PRICE_ASK,
        PRICE_MID_BID_ASK,
        PRICE_BID,
        PRICE_THEORETICAL
    }
}
