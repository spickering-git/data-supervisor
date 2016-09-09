using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCollectionForRealtime
{
    class RealtimeDataObjects
    {
    };

    public class OHLCData
    {
        public DateTime barTime { get; set; }
        public double open { get; set; }
        public double high { get; set; }
        public double low { get; set; }
        public double close { get; set; }
        public int volume { get; set; }
        public int cumulativeVolume { get; set; }

        public bool errorBar { get; set; }
    };

    public class TheoreticalBar
    {
        public DateTime barTime;
        public double price;
    }

    public class Instrument : tblinstrument
    {
        public bool eodAnalysisAtInstrument;
    }

    public class Mongo_OptionSpreadExpression
    {
        [BsonRepresentation(BsonType.ObjectId)]
        public ObjectId Id { get; set; }


        

        //[BsonElement("cqgSymbol")]
        public string cqgSymbol { get; set; }

        public bool stopUpdating = false;

        public int idInstrument;

        public Instrument instrument;

        //public Instrument instrument;

        //this is used for margin calculation from option payoff chart
        //******************************************
        public int optionMonthInt;
        public int optionYear;

        public int futureContractMonthInt;
        public int futureContractYear;
        //******************************************

        public OPTION_EXPRESSION_TYPES optionExpressionType;

        public OPTION_SPREAD_CONTRACT_TYPE callPutOrFuture;

        public char callPutOrFutureChar;


        public int optionId; //only filled if contract is an option
        public int underlyingFutureId;

        public int substituteOptionId; //only filled if contract is an option
        public int substituteUnderlyingFutureId;


        public int idcontract; // futureId; //only filled if contract is a future
        public int substituteFutureId;


        public double strikePrice;

        public double riskFreeRate = 0.01;
        public bool riskFreeRateFilled = false;

        public double yearFraction;

        public DateTime lastTimeUpdated;

        public double minutesSinceLastUpdate = 0;

        public DateTime lastTimeFuturePriceUpdated; //is separate b/c can get time stamp off of historical bars


        public double ask;
        public bool askFilled;

        public double bid;
        public bool bidFilled;

        public double trade;
        public bool tradeFilled;

        public double settlement;
        public bool settlementFilled;
        public bool manuallyFilled;
        public DateTime settlementDateTime;
        public bool settlementIsCurrentDay;

        public double yesterdaySettlement;
        public bool yesterdaySettlementFilled;


        public double defaultBidPriceBeforeTheor;
        //public bool defaultBidPriceBeforeTheorFilled;

        public double defaultAskPriceBeforeTheor;
        //public bool defaultAskPriceBeforeTheorFilled;       

        public double defaultMidPriceBeforeTheor;

        public double defaultPrice;
        public bool defaultPriceFilled;



        public double decisionPrice;
        public DateTime decisionPriceTime;
        public bool decisionPriceFilled = false;

        public double transactionPrice;
        public DateTime transactionPriceTime;
        public bool transactionPriceFilled = false;


        public double impliedVolFromSpan;


        public double theoreticalOptionPrice;

        public double settlementImpliedVol;

        public double impliedVol;

        public bool impliedVolFilled = false; //used for calculating option transaction price

        public double delta;


        public List<OHLCData> futureBarData;
        public List<DateTime> futureBarTimeRef;
        public List<TheoreticalBar> theoreticalOptionDataList;


        public DateTime previousDateTimeBoundaryStart;

        public OHLCData todayTransactionBar;
        public DateTime todayTransactionTimeBoundary;
        public bool reachedTransactionTimeBoundary = false;
        public bool filledAfterTransactionTimeBoundary = false;

        public OHLCData decisionBar;
        public DateTime todayDecisionTime;
        public bool reachedDecisionBar = false;
        public bool reachedBarAfterDecisionBar = false;
        public bool reached1MinAfterDecisionBarUsedForSnapshot = false;

        public bool normalSubscriptionRequest = false;
        public bool substituteSubscriptionRequest = false;

        public bool setSubscriptionLevel = false;
        public bool requestedMinuteBars = false;

        public CQG_REFRESH_STATE totalCalcsRefresh = CQG_REFRESH_STATE.NOTHING;

        public OptionSpreadExpression underlyingFutureExpression; //this field needs to be filled after the database is queried and the expression is an option


        public List<OptionSpreadExpression> optionExpressionsThatUseThisFutureAsUnderlying = new List<OptionSpreadExpression>();
        //if this expression is a future contract this needs to be filled with the options that have this future as an underlying
    }

    public class OptionSpreadExpression : Mongo_OptionSpreadExpression
    {
        public CQG.CQGInstrument cqgInstrument;
        public CQG.CQGTimedBars futureTimedBars;
    }

    //public class OptionSpreadExpression
    //{
    //    public CQG.CQGInstrument cqgInstrument;
    //    public CQG.CQGTimedBars futureTimedBars;

    //    public bool stopUpdating = false;

    //    public int idInstrument;

    //    public Instrument instrument;

    //    public String cqgSymbol;

    //    //public String cqgSubstituteSymbol;


    //    public bool normalSubscriptionRequest = false;
    //    public bool substituteSubscriptionRequest = false;

    //    public bool useSubstituteSymbolAtEOD = false;

    //    //this is used for margin calculation from option payoff chart
    //    //******************************************
    //    public int optionMonthInt;
    //    public int optionYear;

    //    public int futureContractMonthInt;
    //    public int futureContractYear;
    //    //******************************************

    //    public bool setSubscriptionLevel = false;
    //    public bool requestedMinuteBars = false;

    //    public OPTION_EXPRESSION_TYPES optionExpressionType;

    //    public OPTION_SPREAD_CONTRACT_TYPE callPutOrFuture;

    //    public char callPutOrFutureChar;


    //    public int optionId; //only filled if contract is an option
    //    public int underlyingFutureId;

    //    public int substituteOptionId; //only filled if contract is an option
    //    public int substituteUnderlyingFutureId;


    //    public int futureId; //only filled if contract is a future
    //    public int substituteFutureId;


    //    public double strikePrice;

    //    public double riskFreeRate = 0.01;
    //    public bool riskFreeRateFilled = false;

    //    public double yearFraction;

    //    public DateTime lastTimeUpdated;

    //    public double minutesSinceLastUpdate = 0;

    //    public DateTime lastTimeFuturePriceUpdated; //is separate b/c can get time stamp off of historical bars

    //    //public DateTime decisionTime;
    //    //public DateTime transactionTime;
    //    //public DateTime latestDecisionTimeUpdated;
    //    //public DateTime latestTransactionTimeUpdated;

    //    public double ask;
    //    public bool askFilled;

    //    public double bid;
    //    public bool bidFilled;

    //    public double trade;
    //    public bool tradeFilled;

    //    public double settlement;
    //    public bool settlementFilled;
    //    //public bool manuallyFilled;
    //    public DateTime settlementDateTime;
    //    public bool settlementIsCurrentDay;

    //    public double yesterdaySettlement;
    //    public bool yesterdaySettlementFilled;


    //    public double defaultBidPriceBeforeTheor;
    //    //public bool defaultBidPriceBeforeTheorFilled;

    //    public double defaultAskPriceBeforeTheor;
    //    //public bool defaultAskPriceBeforeTheorFilled;        

    //    public double defaultMidPriceBeforeTheor;

    //    public double defaultPrice;
    //    public bool defaultPriceFilled;



    //    public double decisionPrice;
    //    public DateTime decisionPriceTime;
    //    public bool decisionPriceFilled = false;

    //    public double transactionPrice;
    //    public DateTime transactionPriceTime;
    //    public bool transactionPriceFilled = false;


    //    public double impliedVolFromSpan;


    //    public double theoreticalOptionPrice;

    //    public double settlementImpliedVol;

    //    public double impliedVol;

    //    public bool impliedVolFilled = false; //used for calculating option transaction price

    //    public double delta;
    //    //public double gamma;
    //    //public double vega;
    //    //public double theta;



    //    //public List<int> spreadIdx = new List<int>();
    //    //public List<int> legIdx = new List<int>();
    //    //public List<int> rowIdx = new List<int>();


    //    //public List<int> substituteSymbolSpreadIdx = new List<int>();
    //    public List<int> substituteSymbolLegIdx = new List<int>();
    //    public List<int> substituteSymbolRowIdx = new List<int>();

    //    public OptionSpreadExpression mainExpressionSubstitutionUsedFor;


    //    //public List<int> admStrategyIdx = new List<int>();
    //    public List<int> admPositionImportWebIdx = new List<int>();
    //    public List<int> admRowIdx = new List<int>();


    //    public List<OHLCData> futureBarData;
    //    public List<DateTime> futureBarTimeRef;
    //    public List<TheoreticalBar> theoreticalOptionDataList;


    //    public DateTime previousDateTimeBoundaryStart;

    //    public OHLCData todayTransactionBar;
    //    public DateTime todayTransactionTimeBoundary;
    //    public bool reachedTransactionTimeBoundary = false;
    //    public bool filledAfterTransactionTimeBoundary = false;

    //    public OHLCData decisionBar;
    //    public DateTime todayDecisionTime;
    //    public bool reachedDecisionBar = false;
    //    public bool reachedBarAfterDecisionBar = false;
    //    public bool reached1MinAfterDecisionBarUsedForSnapshot = false;

    //    public CQG_REFRESH_STATE guiRefresh = CQG_REFRESH_STATE.NOTHING;
    //    public CQG_REFRESH_STATE totalCalcsRefresh = CQG_REFRESH_STATE.NOTHING;

    //    public OptionSpreadExpression underlyingFutureExpression;

    //    //public List<OptionSpreadExpression> optionExpressionsThatUseThisFutureAsUnderlying;

    //    //Expression List grid
    //    //public int dataGridExpressionListRow;

    //    //************************

    //    public OptionSpreadExpression(OPTION_SPREAD_CONTRACT_TYPE callPutOrFuture,
    //        OPTION_EXPRESSION_TYPES optionExpressionType)
    //    {
    //        this.callPutOrFuture = callPutOrFuture;
    //        this.optionExpressionType = optionExpressionType;

    //        if (optionExpressionType == OPTION_EXPRESSION_TYPES.SPREAD_LEG_PRICE
    //            && callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
    //        {
    //            optionExpressionsThatUseThisFutureAsUnderlying
    //                = new List<OptionSpreadExpression>();
    //        }

    //    }
    //}

    //public class Instrument
    //{
    //    public bool continueRealtimeDataCurrent;

    //    public bool isSpread;

    //    public bool isListedSpread;

    //    public int idxOfInstrumentInList;

    //    public String CQGsymbol;
    //    public String name;
    //    public String exchangeSymbol;
    //    public String optionExchangeSymbol;
    //    public String exchangeSymbolTT;
    //    public String optionExchangeSymbolTT;

    //    public String description;
    //    public String country;
    //    public String currency;

    //    public DateTime dataStart;
    //    public DateTime dataStop;
    //    public DateTime expiration;

    //    public int year;
    //    public char month;

    //    public int idInstrument;
    //    public int idBarInfo;

    //    public int idPortfolioGroup;

    //    public int idInstrumentGroup;
    //    public String instrumentGroup;

    //    public double tickSize;
    //    public double tickDisplay;

    //    /// <summary>
    //    /// The tick display for TT used to multiply, usually just value of 1
    //    /// but for example soybeans price is sent from CQG and SPAN in cents 904.25
    //    /// and optionTickDisplay is 2 makes value 9042
    //    /// TT needs price sent in dollars, so this price is 9.042 this value is 0.001
    //    /// </summary>
    //    //public double tickDisplayTT;

    //    public double tickValue;
    //    public int margin;
    //    public int timeShiftHours;
    //    public double commissionPerContract;


    //    public bool displayInStrategySummary;

    //    public bool continueWithOptimize;

    //    public bool configuredSeries;
    //    public int idcontractseriescfg;

    //    public double optionStrikeIncrement;

    //    public int stopType;

    //    public int limitTickOffset;

    //    public double optionTickSize;
    //    public double optionTickDisplay;

    //    /// <summary>
    //    /// The option tick display for TT used to multiply, usually just value of 1
    //    /// but for example soybeans price is sent from CQG and SPAN in cents 904.25
    //    /// and optionTickDisplay is 2 makes value 9042
    //    /// TT needs price sent in dollars, so this price is 9.042 this value is 0.001
    //    /// </summary>
    //    //public double optionTickDisplayTT;

    //    public double optionTickValue;
    //    public double optionStrikeDisplay;
    //    public double optionStrikeDisplayTT;

    //    public double optionADMStrikeDisplay;
    //    public double admOptionFtpFileStrikeDisplay;

    //    public double admFuturePriceFactor;
    //    public double admOptionPriceFactor;


    //    public double secondaryOptionTickSize;
    //    public double secondaryOptionTickValue;

    //    public double secondaryoptiontickdisplay;

    //    public double secondaryOptionTickSizeRule;

    //    public DateTime customDayBoundaryTime;
    //    public bool useDailyCustomData;
    //    public int decisionOffsetMinutes;

    //    public DateTime optionSpreadStart;

    //    public String admCode;
    //    public String admExchangeCode;

    //    public String exchange;
    //    public String spanExchangeSymbol;
    //    public String spanExchWebAPISymbol;
    //    public String tradingTechnologiesExchange;
    //    public String tradingTechnologiesGateway;

    //    public String spanFutureCode;
    //    public String spanOptionCode;

    //    //public List<TradeCalendarData> tradeCalendarData;
    //    //public TradeCalendarDescription[] tradeCalendarDescription;

    //    public string coreAPImarginId;
    //    public double coreAPIinitialMargin;
    //    public double coreAPImaintenanceMargin;

    //    public string coreAPI_FCM_marginId;
    //    public double coreAPI_FCM_initialMargin;
    //    public double coreAPI_FCM_maintenanceMargin;

    //    public bool substituteSymbolEOD;
    //    public string instrumentSymbolPreEOD;
    //    public string instrumentSymbolEOD;



    //    public DateTime settlementTime;



    //    public bool eodAnalysisAtInstrument;
    //    public DateTime settlementDateTimeMarker;
    //    //public bool reachedSettlementDateTimeMarker;

    //};
}
