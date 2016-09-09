using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CQG;
using System.Collections.Concurrent;
using System.Threading;

namespace DataCollectionForRealtime
{
    class CQGDataManagement
    {
        internal CQGDataManagement(RealtimeDataManagement realtimeDataManagement)
        {
            this.realtimeDataManagement = realtimeDataManagement;

            ThreadPool.QueueUserWorkItem(new WaitCallback(initializeCQGAndCallbacks));
        }

        private DataManagementUtility dataManagementUtility = new DataManagementUtility();

        internal REALTIME_PRICE_FILL_TYPE realtimePriceFillType = REALTIME_PRICE_FILL_TYPE.PRICE_DEFAULT;

        private Thread subscriptionThread;
        private bool subscriptionThreadShouldStop = false;
        private const int SUBSCRIPTION_TIMEDELAY_CONSTANT = 125;

        private RealtimeDataManagement realtimeDataManagement;
        
        private CQG.CQGCEL m_CEL;

        List<OptionSpreadExpression> optionSpreadExpressionList = new List<OptionSpreadExpression>();

        private ConcurrentDictionary<string, OptionSpreadExpression> optionSpreadExpressionHashTable_keySymbol
            = new ConcurrentDictionary<string, OptionSpreadExpression>();

        private ConcurrentDictionary<string, OptionSpreadExpression> optionSpreadExpressionHashTable_keyCQGInId
            = new ConcurrentDictionary<string, OptionSpreadExpression>();

        private ConcurrentDictionary<string, OptionSpreadExpression> optionSpreadExpressionHashTable_keyFullName
            = new ConcurrentDictionary<string, OptionSpreadExpression>();

        internal ConcurrentDictionary<long, Instrument> instrumentHashTable
            = new ConcurrentDictionary<long, Instrument>();

        internal List<Instrument> instrumentList = new List<Instrument>();


        internal void connectCQG()
        {
            if (m_CEL != null)
            {
                m_CEL.Startup();
            }
        }

        internal void shutDownCQGConn()
        {
            if (m_CEL != null)
            {
                if (m_CEL.IsStarted)
                    m_CEL.RemoveAllInstruments();

                //m_CEL.Shutdown();
            }
        }

        internal void initializeCQGAndCallbacks(Object obj)
        {
            try
            {

                m_CEL = new CQG.CQGCEL();

                m_CEL_CELDataConnectionChg(CQG.eConnectionStatus.csConnectionDown);
                //(callsFromCQG,&CallsFromCQG.m_CEL_CELDataConnectionChg);
                m_CEL.DataConnectionStatusChanged += new CQG._ICQGCELEvents_DataConnectionStatusChangedEventHandler(m_CEL_CELDataConnectionChg);


                m_CEL.TimedBarsResolved += new CQG._ICQGCELEvents_TimedBarsResolvedEventHandler(m_CEL_TimedBarResolved);
                m_CEL.TimedBarsAdded += new CQG._ICQGCELEvents_TimedBarsAddedEventHandler(m_CEL_TimedBarsAdded);
                m_CEL.TimedBarsUpdated += new CQG._ICQGCELEvents_TimedBarsUpdatedEventHandler(m_CEL_TimedBarsUpdated);

                //m_CEL.IncorrectSymbol += new _ICQGCELEvents_IncorrectSymbolEventHandler(CEL_IncorrectSymbol);
                m_CEL.InstrumentSubscribed += new _ICQGCELEvents_InstrumentSubscribedEventHandler(m_CEL_InstrumentSubscribed);
                m_CEL.InstrumentChanged += new _ICQGCELEvents_InstrumentChangedEventHandler(m_CEL_InstrumentChanged);

                m_CEL.DataError += new _ICQGCELEvents_DataErrorEventHandler(m_CEL_DataError);

                //m_CEL.APIConfiguration.NewInstrumentMode = true;

                m_CEL.APIConfiguration.ReadyStatusCheck = CQG.eReadyStatusCheck.rscOff;

                m_CEL.APIConfiguration.CollectionsThrowException = false;

                m_CEL.APIConfiguration.TimeZoneCode = CQG.eTimeZone.tzPacific;

                connectCQG();
            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        private void m_CEL_DataError(System.Object cqg_error, System.String error_description)
        {
            try
            {
                if (realtimeDataManagement != null)
                {
                    realtimeDataManagement.updateCQGDataStatus(
                        "CQG ERROR", Color.Yellow, Color.Red);
                }
            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        public void sendSubscribeRequest(bool sendOnlyUnsubscribed)
        {

#if DEBUG
            try
#endif
            {
                subscriptionThread = new Thread(new ParameterizedThreadStart(sendSubscribeRequestRun));
                subscriptionThread.IsBackground = true;
                subscriptionThread.Start(sendOnlyUnsubscribed);

            }
#if DEBUG
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
#endif

        }

        public void sendSubscribeRequestRun(Object obj)
        {
            dataManagementUtility.openThread(null, null);

            try
            {
                //m_CEL.RemoveAllTimedBars();
                //Thread.Sleep(3000);

                if (m_CEL.IsStarted)
                {
                    bool sendOnlyUnsubscribed = (bool)obj;

                    int i = 0;

                    //m_CEL.NewInstrument("C.US.EPN1312600");

                    while (!subscriptionThreadShouldStop && i < optionSpreadExpressionList.Count)
                    {

                        //TSErrorCatch.debugWriteOut("SUBSCRIBE " + optionSpreadExpressionList[i].cqgSymbol);

                        if (sendOnlyUnsubscribed)
                        {


                            if (!optionSpreadExpressionList[i].setSubscriptionLevel)
                            {
                                Thread.Sleep(SUBSCRIPTION_TIMEDELAY_CONSTANT);

                                realtimeDataManagement.updateStatusSubscribeData(
                                    "SUBSCRIBE " + optionSpreadExpressionList[i].cqgSymbol
                                    + " : " + i + " OF " +
                                    optionSpreadExpressionList.Count);

                                m_CEL.NewInstrument(optionSpreadExpressionList[i].cqgSymbol);

                            }

                            if (optionSpreadExpressionList[i].callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                            {
                                if (!optionSpreadExpressionList[i].requestedMinuteBars
                                    && optionSpreadExpressionList[i].normalSubscriptionRequest)
                                {
                                    requestFutureContractTimeBars(optionSpreadExpressionList[i]);
                                }
                            }

                        }
                        else
                        {
                            optionSpreadExpressionList[i].setSubscriptionLevel = false;

                            Thread.Sleep(SUBSCRIPTION_TIMEDELAY_CONSTANT);

                            realtimeDataManagement.updateStatusSubscribeData(
                                    "SUBSCRIBE " + optionSpreadExpressionList[i].cqgSymbol
                                    + " : " + i + " OF " +
                                    optionSpreadExpressionList.Count);

                            m_CEL.NewInstrument(optionSpreadExpressionList[i].cqgSymbol);

                            if (optionSpreadExpressionList[i].callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                            {
                                optionSpreadExpressionList[i].requestedMinuteBars = false;                                

                                if (optionSpreadExpressionList[i].normalSubscriptionRequest)
                                {
                                    requestFutureContractTimeBars(optionSpreadExpressionList[i]);
                                }

                            }
                        }



                        i++;
                    }

                    Thread.Sleep(SUBSCRIPTION_TIMEDELAY_CONSTANT);

                    realtimeDataManagement.updateStatusSubscribeData("");

                }
            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }

            dataManagementUtility.closeThread(null, null);
        }

        public void requestFutureContractTimeBars(OptionSpreadExpression optionSpreadExpression)
        {
            try
            {

                CQGTimedBarsRequest timedBarsRequest = m_CEL.CreateTimedBarsRequest();

                timedBarsRequest.Symbol = optionSpreadExpression.cqgSymbol;

                timedBarsRequest.SessionsFilter = 31;

                timedBarsRequest.IntradayPeriod = 1;

                timedBarsRequest.Continuation = CQG.eTimeSeriesContinuationType.tsctNoContinuation;
                //do not want continuation bars

                DateTime rangeStart = optionSpreadExpression.previousDateTimeBoundaryStart;

                DateTime rangeEnd = m_CEL.Environment.LineTime;

                timedBarsRequest.RangeStart = rangeStart;

                timedBarsRequest.RangeEnd = rangeEnd;

                timedBarsRequest.IncludeEnd = true;

                timedBarsRequest.UpdatesEnabled = true;

                optionSpreadExpression.futureTimedBars = m_CEL.RequestTimedBars(timedBarsRequest);


                optionSpreadExpressionHashTable_keyCQGInId.AddOrUpdate(
                    optionSpreadExpression.futureTimedBars.Id,
                    optionSpreadExpression, (oldKey, oldValue) => optionSpreadExpression);
            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        

        private void m_CEL_CELDataConnectionChg(CQG.eConnectionStatus new_status)
        {
            StringBuilder connStatusString = new StringBuilder();
            StringBuilder connStatusShortString = new StringBuilder();
            Color connColor = Color.Red;

            try
            {
                if (m_CEL.IsStarted)
                {
                    connStatusString.Append("CQG API:");
                    connStatusString.Append(m_CEL.Environment.CELVersion);
                    connStatusShortString.Append("CQG:");

                    if (new_status != CQG.eConnectionStatus.csConnectionUp)
                    {
                        if (new_status == CQG.eConnectionStatus.csConnectionDelayed)
                        {
                            connColor = Color.BlanchedAlmond;
                            connStatusString.Append(" - CONNECTION IS DELAYED");
                            connStatusShortString.Append("DELAYED");
                        }
                        else
                        {
                            connStatusString.Append(" - CONNECTION IS DOWN");
                            connStatusShortString.Append("DOWN");
                        }
                    }
                    else
                    {
                        connColor = Color.LawnGreen;
                        connStatusString.Append(" - CONNECTION IS UP");
                        connStatusShortString.Append("UP");
                    }
                }
                else
                {
                    connStatusString.Append("WAITING FOR API CONNECTION");

                    connStatusShortString.Append("WAITING");
                }

                if (realtimeDataManagement != null)
                {
                    realtimeDataManagement.updateConnectionStatus(
                        connStatusString.ToString(), connColor);
                }
            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        private void m_CEL_TimedBarResolved(CQG.CQGTimedBars cqg_TimedBarsIn, CQGError cqg_error)
        {
            //Debug.WriteLine("m_CEL_ExpressionResolved" + cqg_expression.Count);
            try
            {
                //TSErrorCatch.debugWriteOut(cqg_TimedBarsIn.Id);

                if (cqg_error == null)
                {
                    //int expressionCounter = 0;

                    if (optionSpreadExpressionHashTable_keyCQGInId.ContainsKey(cqg_TimedBarsIn.Id))
                    {

                        OptionSpreadExpression optionSpreadExpression
                            = optionSpreadExpressionHashTable_keyCQGInId[cqg_TimedBarsIn.Id];

                        Instrument instrument = null;

                        if (instrumentHashTable.ContainsKey(optionSpreadExpression.idInstrument))
                        {
                            instrument = instrumentHashTable[optionSpreadExpression.idInstrument];
                        }

                        //while (expressionCounter < optionSpreadExpressionList.Count)
                        //{
                        if (optionSpreadExpression.stopUpdating
                            && optionSpreadExpression.futureTimedBars != null)
                        {
                            optionSpreadExpression.requestedMinuteBars = true;

                            optionSpreadExpression.futureBarData
                                = new List<OHLCData>(DataCollectionConstants.MINUTES_IN_DAY);

                            //OHLCData ohlcData = new OHLCData();

                            //optionSpreadExpressionList[expressionCounter].futureBarData.Add(ohlcData);

                            //TSErrorCatch.debugWriteOut("resolved " + cqg_TimedBarsIn.Id
                            //    + "  " + cqg_TimedBarsIn.Count);

                            if (cqg_TimedBarsIn.Count > 0)
                            {

                                int timeBarsIn_CurrentDay_TransactionIdx = cqg_TimedBarsIn.Count - 1;

                                DateTime currentDay = DateTime.Now;

                                if (instrument != null)
                                {

                                    //changed decision and transaction date to the modelDateTime rather than
                                    //the date of the latest bar. Because it was failing on instruments like cattle
                                    //and hogs that don't have data
                                    optionSpreadExpression.todayTransactionTimeBoundary
                                        = currentDay.Date
                                        .AddHours(
                                           instrument.customdayboundarytime.Hour)
                                        .AddMinutes(
                                            instrument.customdayboundarytime.Minute);

                                    optionSpreadExpression.todayDecisionTime
                                        = currentDay.Date
                                        .AddHours(
                                            instrument.customdayboundarytime.Hour)
                                        .AddMinutes(
                                            instrument.customdayboundarytime.Minute
                                            - instrument.decisionoffsetminutes);
                                }

                                //TSErrorCatch.debugWriteOut(
                                //    optionSpreadExpressionList[expressionCounter].cqgSymbol + " " +
                                //    optionSpreadExpressionList[expressionCounter].todayDecisionTime)

                                //optionSpreadExpressionList[expressionCounter].settlementDateTimeMarker
                                //    = optionSpreadManager.initializationParmsSaved.modelDateTime.Date
                                //    .AddHours(
                                //        optionSpreadExpressionList[expressionCounter].instrument.settlementTime.Hour)
                                //    .AddMinutes(
                                //        optionSpreadExpressionList[expressionCounter].instrument.settlementTime.Minute
                                //        + 15);

                                //TSErrorCatch.debugWriteOut(
                                //    optionSpreadExpressionList[expressionCounter].instrument.CQGsymbol + "  " +
                                //    optionSpreadExpressionList[expressionCounter].settlementDateTimeMarker);


                                int timeBarsIn_Counter = 0;
                                int cumVolume = 0;


                                while (timeBarsIn_Counter < cqg_TimedBarsIn.Count)
                                {




                                    if (cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp
                                        .CompareTo(optionSpreadExpression.previousDateTimeBoundaryStart) >= 0)
                                    //&& cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp
                                    //.CompareTo(optionSpreadExpressionList[expressionCounter].todayTransactionTimeBoundary) <= 0)
                                    {
                                        bool error = false;

                                        OHLCData ohlcData = new OHLCData();

                                        optionSpreadExpression.futureBarData.Add(ohlcData);

                                        ohlcData.barTime = cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp;

                                        int volume = 0;

                                        if (cqg_TimedBarsIn[timeBarsIn_Counter].ActualVolume
                                            != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                        {
                                            volume = cqg_TimedBarsIn[timeBarsIn_Counter].ActualVolume;
                                        }
                                        else
                                        {
                                            error = true;
                                        }

                                        //if (timeBarsIn_Counter == 0)
                                        //{
                                        //    TSErrorCatch.debugWriteOut(
                                        //        optionSpreadExpressionList[expressionCounter].cqgSymbol +
                                        //        " VOL BAR " + cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp);
                                        //}
                                        //else if (timeBarsIn_Counter == cqg_TimedBarsIn.Count - 1)
                                        //{
                                        //    TSErrorCatch.debugWriteOut(
                                        //        optionSpreadExpressionList[expressionCounter].cqgSymbol +
                                        //        " VOL BAR " + cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp);
                                        //}

                                        cumVolume += volume;

                                        ohlcData.cumulativeVolume = cumVolume;

                                        ohlcData.volume = volume;



                                        if (cqg_TimedBarsIn[timeBarsIn_Counter].Open
                                            != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                        {
                                            ohlcData.open = cqg_TimedBarsIn[timeBarsIn_Counter].Open;
                                        }
                                        else
                                        {
                                            error = true;
                                        }

                                        if (cqg_TimedBarsIn[timeBarsIn_Counter].High
                                            != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                        {
                                            ohlcData.high = cqg_TimedBarsIn[timeBarsIn_Counter].High;
                                        }
                                        else
                                        {
                                            error = true;
                                        }

                                        if (cqg_TimedBarsIn[timeBarsIn_Counter].Low
                                            != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                        {
                                            ohlcData.low = cqg_TimedBarsIn[timeBarsIn_Counter].Low;
                                        }
                                        else
                                        {
                                            error = true;
                                        }


                                        if (cqg_TimedBarsIn[timeBarsIn_Counter].Close
                                        != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                        {
                                            ohlcData.close = cqg_TimedBarsIn[timeBarsIn_Counter].Close;
                                        }
                                        else
                                        {
                                            error = true;
                                        }

                                        ohlcData.errorBar = error;



                                        if (!error
                                            && !optionSpreadExpression.reachedTransactionTimeBoundary
                                            && cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp
                                            .CompareTo(optionSpreadExpression.todayTransactionTimeBoundary) <= 0)
                                        {
                                            optionSpreadExpression.todayTransactionBar = ohlcData;
                                        }

                                        if (!error
                                            && !optionSpreadExpression.reachedTransactionTimeBoundary
                                            && cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp
                                            .CompareTo(optionSpreadExpression.todayTransactionTimeBoundary) >= 0)
                                        {
                                            optionSpreadExpression.reachedTransactionTimeBoundary = true;
                                        }

                                        if (!error
                                            && !optionSpreadExpression.reachedDecisionBar
                                            && cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp
                                            .CompareTo(optionSpreadExpression.todayDecisionTime) <= 0)
                                        {
                                            optionSpreadExpression.decisionBar = ohlcData;
                                        }

                                        if (!error
                                            && !optionSpreadExpression.reachedDecisionBar
                                            && cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp
                                            .CompareTo(optionSpreadExpression.todayDecisionTime) >= 0)
                                        {
                                            optionSpreadExpression.reachedDecisionBar = true;


                                        }

                                        if (!error
                                            && !optionSpreadExpression.reachedBarAfterDecisionBar
                                            && cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp
                                            .CompareTo(optionSpreadExpression.todayDecisionTime) > 0)
                                        {
                                            optionSpreadExpression.reachedBarAfterDecisionBar = true;
                                        }

                                        if (!error
                                            && !optionSpreadExpression.reached1MinAfterDecisionBarUsedForSnapshot
                                            && cqg_TimedBarsIn[timeBarsIn_Counter].Timestamp
                                            .CompareTo(optionSpreadExpression.todayDecisionTime.AddMinutes(1)) > 0)
                                        {
                                            optionSpreadExpression.reached1MinAfterDecisionBarUsedForSnapshot = true;
                                        }


                                    }

                                    timeBarsIn_Counter++;
                                }

                                int backTimeCounter = cqg_TimedBarsIn.Count - 1;

                                while (backTimeCounter >= 0)
                                {
                                    if (cqg_TimedBarsIn[backTimeCounter].Close
                                    != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                    {
                                        optionSpreadExpression.lastTimeFuturePriceUpdated =
                                            cqg_TimedBarsIn[backTimeCounter].Timestamp;

                                        optionSpreadExpression.trade =
                                            cqg_TimedBarsIn[backTimeCounter].Close;

                                        optionSpreadExpression.tradeFilled = true;

                                        break;
                                    }

                                    backTimeCounter--;
                                }



                                fillDefaultMidPrice(optionSpreadExpression);

                                manageExpressionPriceCalcs(optionSpreadExpression);

                            }

                            //break;  //while (expressionCounter < optionSpreadExpressionList.Count)

                        }

                        //expressionCounter++;

                    }


                }
                else
                {
                    realtimeDataManagement.updateCQGDataStatus("CQG ERROR", Color.Yellow, Color.Red);
                }
            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        private void m_CEL_TimedBarsAdded(CQG.CQGTimedBars cqg_TimedBarsIn)
        {
            try
            {
                //TSErrorCatch.debugWriteOut(cqg_TimedBarsIn.Id);

                {

                    //if (optionSpreadExpressionFutureTimedBarsListIdx.Count > 0
                    //    &&
                    //    optionSpreadExpressionFutureTimedBarsListIdx.ContainsKey(cqg_TimedBarsIn.Id))
                    
                    if (optionSpreadExpressionHashTable_keyCQGInId.ContainsKey(cqg_TimedBarsIn.Id))
                    {

                        OptionSpreadExpression optionSpreadExpression = optionSpreadExpressionHashTable_keyCQGInId[cqg_TimedBarsIn.Id];

                        //while (expressionCounter < optionSpreadExpressionList.Count)
                        //{
                        if (!optionSpreadExpression.stopUpdating
                            && optionSpreadExpression.futureTimedBars != null)

                        //while (expressionCounter < optionSpreadExpressionList.Count)
                        //{
                        //    if (!optionSpreadExpressionList[expressionCounter].stopUpdating
                        //        && optionSpreadExpressionList[expressionCounter].futureTimedBars != null
                        //        && cqg_TimedBarsIn.Id.CompareTo(optionSpreadExpressionList[expressionCounter].futureTimedBars.Id) == 0)
                        //&& !optionSpreadExpressionList[expressionCounter].requestedMinuteBars)
                        {
                            int lastTimdBarsInIdx = cqg_TimedBarsIn.Count - 1;

                            while (lastTimdBarsInIdx >= 0)
                            {
                                if (cqg_TimedBarsIn[lastTimdBarsInIdx].Timestamp.CompareTo(
                                    optionSpreadExpression.futureBarData.Last().barTime) <= 0)
                                {
                                    lastTimdBarsInIdx++;
                                    break;
                                }

                                lastTimdBarsInIdx--;
                            }

                            if (lastTimdBarsInIdx < 0)
                            {
                                lastTimdBarsInIdx = 0;
                            }

                            while (lastTimdBarsInIdx < cqg_TimedBarsIn.Count)
                            {

                                bool error = false;


                                OHLCData ohlcData = new OHLCData();

                                optionSpreadExpression.futureBarData.Add(ohlcData);

                                ohlcData.barTime = cqg_TimedBarsIn[lastTimdBarsInIdx].Timestamp;

                                int volume = 0;

                                if (cqg_TimedBarsIn[lastTimdBarsInIdx].ActualVolume
                                    != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                {
                                    volume = cqg_TimedBarsIn[lastTimdBarsInIdx].ActualVolume;
                                }
                                else
                                {
                                    error = true;
                                }



                                ohlcData.volume = volume;


                                int cumVolume = 0;

                                int cumCounter = 0;

                                foreach (OHLCData ohlcDataCumvol in optionSpreadExpression.futureBarData)
                                {
                                    if (ohlcDataCumvol.barTime.CompareTo(
                                        optionSpreadExpression.previousDateTimeBoundaryStart) >= 0)
                                    {
                                        //if (cumCounter == 0)
                                        //{
                                        //    TSErrorCatch.debugWriteOut(
                                        //        optionSpreadExpressionList[expressionCounter].cqgSymbol +
                                        //        " VOL BAR " + ohlcDataCumvol.barTime);
                                        //}
                                        //else if (cumCounter == 
                                        //    optionSpreadExpressionList[expressionCounter].futureBarData.Count - 1)
                                        //{
                                        //    TSErrorCatch.debugWriteOut(
                                        //        optionSpreadExpressionList[expressionCounter].cqgSymbol +
                                        //        " VOL BAR " + ohlcDataCumvol.barTime);
                                        //}

                                        cumCounter++;

                                        cumVolume += ohlcDataCumvol.volume;

                                        ohlcDataCumvol.cumulativeVolume = cumVolume;
                                    }
                                }


                                //int timeBarsIn_Counter = 0;

                                //while (timeBarsIn_Counter < cqg_TimedBarsIn.Count)
                                //{


                                //    if (cqg_TimedBarsIn[timeBarsIn_Counter].ActualVolume
                                //                != -TradingSystemConstants.CQG_DATA_ERROR_CODE)
                                //    {
                                //        //volume = cqg_TimedBarsIn[timeBarsIn_Counter].ActualVolume;

                                //        cumVolume += cqg_TimedBarsIn[timeBarsIn_Counter].ActualVolume;
                                //    }

                                //    timeBarsIn_Counter++;

                                //}

                                ohlcData.cumulativeVolume = cumVolume;


                                if (cqg_TimedBarsIn[lastTimdBarsInIdx].Open
                                    != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                {
                                    ohlcData.open = cqg_TimedBarsIn[lastTimdBarsInIdx].Open;
                                }
                                else
                                {
                                    error = true;
                                }

                                if (cqg_TimedBarsIn[lastTimdBarsInIdx].High
                                    != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                {
                                    ohlcData.high = cqg_TimedBarsIn[lastTimdBarsInIdx].High;
                                }
                                else
                                {
                                    error = true;
                                }

                                if (cqg_TimedBarsIn[lastTimdBarsInIdx].Low
                                    != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                {
                                    ohlcData.low = cqg_TimedBarsIn[lastTimdBarsInIdx].Low;
                                }
                                else
                                {
                                    error = true;
                                }

                                if (cqg_TimedBarsIn[lastTimdBarsInIdx].Close
                                    != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                {
                                    ohlcData.close = cqg_TimedBarsIn[lastTimdBarsInIdx].Close;
                                }
                                else
                                {
                                    error = true;
                                }

                                ohlcData.errorBar = error;


                                //OHLCData futureBarData = optionSpreadExpressionList[expressionCounter].futureBarData[0];

                                //futureBarData.barTime = ohlcData.barTime;

                                //futureBarData.open = ohlcData.open;

                                //futureBarData.high = ohlcData.high;

                                //futureBarData.low = ohlcData.low;

                                //futureBarData.close = ohlcData.close;

                                //futureBarData.volume = ohlcData.volume;

                                //futureBarData.cumulativeVolume = ohlcData.cumulativeVolume;

                                //futureBarData.errorBar = ohlcData.errorBar;


                                if (!error
                                    && !optionSpreadExpression.reachedTransactionTimeBoundary
                                    && ohlcData.barTime
                                    .CompareTo(optionSpreadExpression.todayTransactionTimeBoundary) <= 0)
                                {
                                    optionSpreadExpression.todayTransactionBar = ohlcData;
                                }

                                if (!error
                                    && !optionSpreadExpression.reachedTransactionTimeBoundary
                                    && ohlcData.barTime
                                    .CompareTo(optionSpreadExpression.todayTransactionTimeBoundary) >= 0)
                                {
                                    optionSpreadExpression.reachedTransactionTimeBoundary = true;
                                }

                                if (!error
                                    && !optionSpreadExpression.reachedDecisionBar
                                    && ohlcData.barTime
                                    .CompareTo(optionSpreadExpression.todayDecisionTime) <= 0)
                                {
                                    optionSpreadExpression.decisionBar = ohlcData;
                                }

                                if (!error
                                    && !optionSpreadExpression.reachedDecisionBar
                                    && ohlcData.barTime
                                    .CompareTo(optionSpreadExpression.todayDecisionTime) >= 0)
                                {
                                    optionSpreadExpression.reachedDecisionBar = true;
                                }

                                if (!error
                                    && !optionSpreadExpression.reachedBarAfterDecisionBar
                                    && ohlcData.barTime
                                    .CompareTo(optionSpreadExpression.todayDecisionTime) > 0)
                                {
                                    optionSpreadExpression.reachedBarAfterDecisionBar = true;
                                }

                                if (!error
                                    && !optionSpreadExpression.reached1MinAfterDecisionBarUsedForSnapshot
                                    && ohlcData.barTime
                                    .CompareTo(optionSpreadExpression.todayDecisionTime.AddMinutes(1)) > 0)
                                {
                                    optionSpreadExpression.reached1MinAfterDecisionBarUsedForSnapshot = true;
                                }



                                if (cqg_TimedBarsIn[lastTimdBarsInIdx].Close
                                        != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                                {
                                    optionSpreadExpression.lastTimeFuturePriceUpdated =
                                                cqg_TimedBarsIn[lastTimdBarsInIdx].Timestamp;

                                    optionSpreadExpression.trade =
                                            cqg_TimedBarsIn[lastTimdBarsInIdx].Close;

                                    optionSpreadExpression.tradeFilled = true;

                                    fillDefaultMidPrice(optionSpreadExpression);

                                    manageExpressionPriceCalcs(optionSpreadExpression);
                                }

                                lastTimdBarsInIdx++;

                            }



                        }

                    }



                }

            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        private void m_CEL_TimedBarsUpdated(CQG.CQGTimedBars cqg_TimedBarsIn, int index)
        {
            //Debug.WriteLine("m_CEL_ExpressionResolved" + cqg_expression.Count);
            try
            {

                if (optionSpreadExpressionHashTable_keyCQGInId.ContainsKey(cqg_TimedBarsIn.Id))
                {

                    OptionSpreadExpression optionSpreadExpression
                            = optionSpreadExpressionHashTable_keyCQGInId[cqg_TimedBarsIn.Id];


                    if (!optionSpreadExpression.stopUpdating
                        && optionSpreadExpression.futureTimedBars != null
                        && optionSpreadExpression.futureBarData != null
                        && optionSpreadExpression.futureBarData.Count > 0)
                    {
                        int futureBarDataCounter = optionSpreadExpression.futureBarData.Count - 1;

                        bool foundBar = false;

                        while (futureBarDataCounter >= 0)
                        {
                            if (cqg_TimedBarsIn[index].Timestamp.CompareTo(
                                optionSpreadExpression.futureBarData[futureBarDataCounter].barTime) == 0)
                            {
                                foundBar = true;
                                break;
                            }

                            futureBarDataCounter--;
                        }

                        if (foundBar)
                        {

                            //***********************************


                            OHLCData ohlcData = optionSpreadExpression.futureBarData[futureBarDataCounter];

                            bool error = false;

                            int volume = 0;

                            if (cqg_TimedBarsIn[index].ActualVolume
                                != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                            {
                                volume = cqg_TimedBarsIn[index].ActualVolume;
                            }
                            else
                            {
                                error = true;
                            }

                            ohlcData.volume = volume;

                            if (cqg_TimedBarsIn[index].Open
                                != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                            {
                                ohlcData.open = cqg_TimedBarsIn[index].Open;
                            }
                            else
                            {
                                error = true;
                            }

                            //**********************************************


                            if (cqg_TimedBarsIn[index].High
                                != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                            {
                                ohlcData.high = cqg_TimedBarsIn[index].High;
                            }
                            else
                            {
                                error = true;
                            }

                            if (cqg_TimedBarsIn[index].Low
                                != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                            {
                                ohlcData.low = cqg_TimedBarsIn[index].Low;
                            }
                            else
                            {
                                error = true;
                            }

                            if (cqg_TimedBarsIn[index].Close
                                != -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                            {
                                ohlcData.close = cqg_TimedBarsIn[index].Close;
                            }
                            else
                            {
                                error = true;
                            }

                            ohlcData.errorBar = error;


                            //**********************************************




                            if (!error
                                && !optionSpreadExpression.reachedTransactionTimeBoundary
                                && ohlcData.barTime
                                .CompareTo(optionSpreadExpression.todayTransactionTimeBoundary) <= 0)
                            {
                                optionSpreadExpression.todayTransactionBar = ohlcData;
                            }

                            if (!error
                                && !optionSpreadExpression.reachedTransactionTimeBoundary
                                && ohlcData.barTime
                                .CompareTo(optionSpreadExpression.todayTransactionTimeBoundary) >= 0)
                            {
                                optionSpreadExpression.reachedTransactionTimeBoundary = true;
                            }

                            if (!error
                                && !optionSpreadExpression.reachedDecisionBar
                                && ohlcData.barTime
                                .CompareTo(optionSpreadExpression.todayDecisionTime) <= 0)
                            {
                                optionSpreadExpression.decisionBar = ohlcData;
                            }

                            if (!error
                                && !optionSpreadExpression.reachedDecisionBar
                                && ohlcData.barTime
                                .CompareTo(optionSpreadExpression.todayDecisionTime) >= 0)
                            {
                                optionSpreadExpression.reachedDecisionBar = true;
                            }

                            if (!error
                                && !optionSpreadExpression.reachedBarAfterDecisionBar
                                && ohlcData.barTime
                                .CompareTo(optionSpreadExpression.todayDecisionTime) > 0)
                            {
                                optionSpreadExpression.reachedBarAfterDecisionBar = true;
                            }

                            if (!error
                                && !optionSpreadExpression.reached1MinAfterDecisionBarUsedForSnapshot
                                && ohlcData.barTime
                                .CompareTo(optionSpreadExpression.todayDecisionTime.AddMinutes(1)) > 0)
                            {
                                optionSpreadExpression.reached1MinAfterDecisionBarUsedForSnapshot = true;
                            }


                            if (!optionSpreadExpression.futureBarData.Last().errorBar)
                            //!= -DataCollectionConstants.CQG_DATA_ERROR_CODE)
                            {
                                optionSpreadExpression.lastTimeFuturePriceUpdated =
                                    optionSpreadExpression.futureBarData.Last().barTime;

                                optionSpreadExpression.trade =
                                    optionSpreadExpression.futureBarData.Last().close;

                                fillDefaultMidPrice(optionSpreadExpression);

                                manageExpressionPriceCalcs(optionSpreadExpression);

                            }

                        }

                        //break;  //while (expressionCounter < optionSpreadExpressionList.Count)

                    }

                    //expressionCounter++;
                }

            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        private void m_CEL_InstrumentSubscribed(String symbol, CQGInstrument cqgInstrument)
        {
            try
            {
                realtimeDataManagement.updateCQGDataStatus("CQG GOOD", Color.Black, Color.LawnGreen);


                if (optionSpreadExpressionHashTable_keySymbol.ContainsKey(symbol))
                {

                    OptionSpreadExpression optionSpreadExpression = 
                        optionSpreadExpressionHashTable_keySymbol[symbol];

                    //while (expressionCounter < optionSpreadExpressionList.Count)
                    //{
                    if (!optionSpreadExpression.stopUpdating
                        //&& symbol.CompareTo(optionSpreadExpressionList[expressionCounter].cqgSymbol) == 0
                        && !optionSpreadExpression.setSubscriptionLevel)
                    {
                        optionSpreadExpression.setSubscriptionLevel = true;

                        optionSpreadExpression.cqgInstrument = cqgInstrument;


                        //int idx = expressionCounter;

                        //optionSpreadExpressionListHashTableIdx.AddOrUpdate(
                        //        cqgInstrument.FullName, idx,
                        //        (oldKey, oldValue) => idx);

                        optionSpreadExpressionHashTable_keyFullName.AddOrUpdate(
                                cqgInstrument.FullName, optionSpreadExpression,
                                (oldKey, oldValue) => optionSpreadExpression);

                        //if (cqgInstrument.FullName.CompareTo("P.US.EU6J1511100") == 0)
                        //{
                        //    Console.WriteLine(cqgInstrument.FullName);
                        //}

                        fillPricesFromQuote(optionSpreadExpression,
                            optionSpreadExpression.cqgInstrument.Quotes);

                        //if is an option (not a future)
                        if (optionSpreadExpression.callPutOrFuture !=
                                OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                        {
                            fillDefaultMidPrice(optionSpreadExpression);

                            manageExpressionPriceCalcs(optionSpreadExpression);
                        }


                        if (optionSpreadExpression.optionExpressionType
                            == OPTION_EXPRESSION_TYPES.SPREAD_LEG_PRICE)
                        {
                            //optionSpreadExpression.updatedDataFromCQG = true;

                            if (optionSpreadExpression.callPutOrFuture ==
                                    OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                            {
                                optionSpreadExpression.cqgInstrument.DataSubscriptionLevel
                                    = eDataSubscriptionLevel.dsQuotes;
                            }
                            else
                            {
                                optionSpreadExpression.cqgInstrument.DataSubscriptionLevel
                                    = eDataSubscriptionLevel.dsQuotesAndBBA;

                                optionSpreadExpression.cqgInstrument.BBAType
                                     = eDOMandBBAType.dbtCombined;
                            }

                        }
                        else
                        {
                            //TSErrorCatch.debugWriteOut(symbol + " interest rate no update");

                            //set updates for interest rate to stop
                            optionSpreadExpression.cqgInstrument.DataSubscriptionLevel
                                = eDataSubscriptionLevel.dsNone;

                        }

                        //break;
                    }

                    //expressionCounter++;
                }

            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        private void m_CEL_InstrumentChanged(CQGInstrument cqgInstrument,
                                 CQGQuotes quotes,
                                 CQGInstrumentProperties props)
        {
            try
            {

                if (optionSpreadExpressionHashTable_keyFullName.ContainsKey(cqgInstrument.FullName))
                {

                    //optionSpreadExpressionCheckSubscribedListIdx

                    OptionSpreadExpression optionSpreadExpression 
                        = optionSpreadExpressionHashTable_keyFullName[cqgInstrument.FullName];

                    if (optionSpreadExpression != null
                        && !optionSpreadExpression.stopUpdating
                        && optionSpreadExpression.cqgInstrument != null

                        && cqgInstrument.CEL != null)
                     {
                        //optionSpreadExpressionList[expressionCounter].cqgInstrument = cqgInstrument;

                            CQGQuote quoteAsk = quotes[eQuoteType.qtAsk];
                            CQGQuote quoteBid = quotes[eQuoteType.qtBid];
                            CQGQuote quoteTrade = quotes[eQuoteType.qtTrade];
                            CQGQuote quoteSettlement = quotes[eQuoteType.qtSettlement];
                            CQGQuote quoteYestSettlement = quotes[eQuoteType.qtYesterdaySettlement];

                            if ((quoteAsk != null)
                                || (quoteBid != null)
                                || (quoteTrade != null)
                                || (quoteSettlement != null)
                                || (quoteYestSettlement != null))
                            {
                                //                                 if (optionSpreadExpressionList[expressionCounter].callPutOrFuture
                                //                                     != (int)OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                                //                                 {
                                //                                     TSErrorCatch.debugWriteOut(
                                //                                         cqgInstrument.FullName + "  " +
                                //                                         optionSpreadExpressionList[expressionCounter].cqgInstrument.FullName + "  " +
                                //                                         optionSpreadExpressionList[expressionCounter].cqgSymbol + "  " +
                                //                                         "ASK " + ((quoteAsk != null && quoteAsk.IsValid) ? quoteAsk.Price.ToString() : "blank") + " " +
                                //                                         "BID " + ((quoteBid != null && quoteBid.IsValid) ? quoteBid.Price.ToString() : "blank") + " " +
                                //                                         "TRADE " + ((quoteTrade != null && quoteTrade.IsValid) ? quoteTrade.Price.ToString() : "blank") + " " +
                                //                                         "SETTL " + ((quoteSettlement != null && quoteSettlement.IsValid) ? quoteSettlement.Price.ToString() : "blank") + " " +
                                //                                         "YEST " + ((quoteYestSettlement != null && quoteYestSettlement.IsValid) ? quoteYestSettlement.Price.ToString() : "blank") + " "
                                //                                         );
                                //                                 }

                                //                                 quoteValue =
                                //                                     optionSpreadExpressionList[expressionCounter].cqgInstrument.ToDisplayPrice(quote.Price);

                                fillPricesFromQuote(optionSpreadExpression,
                                    optionSpreadExpression.cqgInstrument.Quotes);

                                if (optionSpreadExpression.callPutOrFuture !=
                                        OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                                {
                                    fillDefaultMidPrice(optionSpreadExpression);

                                    manageExpressionPriceCalcs(optionSpreadExpression);
                                }

                            }
                        

                        //break;
                    }

                }
            }
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
        }

        private void fillDefaultMidPrice(OptionSpreadExpression optionSpreadExpression)  //, Instrument instrument)
        {

            double defaultPrice = 0;

            if (optionSpreadExpression.optionExpressionType == OPTION_EXPRESSION_TYPES.OPTION_EXPRESSION_RISK_FREE_RATE)
            {
                if (optionSpreadExpression.tradeFilled)
                {
                    defaultPrice = optionSpreadExpression.trade;
                }
                else if (optionSpreadExpression.settlementFilled)
                {
                    defaultPrice = optionSpreadExpression.settlement;
                }
                else if (optionSpreadExpression.yesterdaySettlementFilled)
                {
                    defaultPrice = optionSpreadExpression.yesterdaySettlement;
                }

                if (defaultPrice == 0)
                {
                    defaultPrice = 0.01;
                }

                defaultPrice = defaultPrice == 0 ? 0 :
                    ((int)((100 - defaultPrice + DataCollectionConstants.EPSILON) * 1000) / 100000.0);

                optionSpreadExpression.riskFreeRate = defaultPrice;
            }
            else
            {

                if (optionSpreadExpression.callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                {


                    optionSpreadExpression.lastTimeUpdated = optionSpreadExpression.lastTimeFuturePriceUpdated;

                    TimeSpan span = DateTime.Now - optionSpreadExpression.lastTimeUpdated;

                    optionSpreadExpression.minutesSinceLastUpdate = span.TotalMinutes;

                    if (optionSpreadExpression.tradeFilled)
                    {
                        defaultPrice = optionSpreadExpression.trade;
                    }
                    else if (optionSpreadExpression.settlementFilled)
                    {
                        defaultPrice = optionSpreadExpression.settlement;
                    }
                    else if (optionSpreadExpression.yesterdaySettlementFilled)
                    {
                        defaultPrice = optionSpreadExpression.yesterdaySettlement;
                    }

                    if (defaultPrice == 0)
                    {
                        defaultPrice = DataCollectionConstants.OPTION_ZERO_PRICE;
                    }


                    //can set default price for futures here b/c no further price possibilities for future;
                    optionSpreadExpression.defaultPrice = defaultPrice;

                    optionSpreadExpression.defaultPriceFilled = true;

                }
                else //CHANGED DEC 30 2015 
                //to else from else if
                //if(!optionSpreadExpression.instrument.eodAnalysisAtInstrument)
                {


                    optionSpreadExpression.lastTimeUpdated = DateTime.Now;

                    optionSpreadExpression.minutesSinceLastUpdate = 0;

                    if (optionSpreadExpression.bidFilled
                    && optionSpreadExpression.askFilled)
                    {


                        defaultPrice = (optionSpreadExpression.bid + optionSpreadExpression.ask) / 2;

                        //rounding to nearest tick;
                        if (optionSpreadExpression.instrument != null)
                        {

                            double optionTickSize = DataManagementUtility.chooseOptionTickSize(defaultPrice,
                                optionSpreadExpression.instrument.optionticksize,
                                optionSpreadExpression.instrument.secondaryoptionticksize,
                                optionSpreadExpression.instrument.secondaryoptionticksizerule);

                            defaultPrice = ((int)((defaultPrice + optionTickSize / 2) /
                                optionTickSize)) * optionTickSize;

                        }


                    }
                    else if (optionSpreadExpression.askFilled)
                    {
                        defaultPrice = optionSpreadExpression.ask;
                    }
                    else if (optionSpreadExpression.bidFilled)
                    {
                        defaultPrice = optionSpreadExpression.bid;
                    }

                    if (defaultPrice == 0)
                    {
                        if (optionSpreadExpression.instrument != null)
                        {
                            defaultPrice = optionSpreadExpression.instrument.optionticksize;  //OptionConstants.OPTION_ZERO_PRICE;
                        }
                    }

                    optionSpreadExpression.defaultMidPriceBeforeTheor = defaultPrice;

                    if (optionSpreadExpression.askFilled)
                    {
                        optionSpreadExpression.defaultAskPriceBeforeTheor = optionSpreadExpression.ask;
                    }
                    else
                    {
                        optionSpreadExpression.defaultAskPriceBeforeTheor = defaultPrice;
                    }

                    if (optionSpreadExpression.bidFilled)
                    {
                        optionSpreadExpression.defaultBidPriceBeforeTheor = optionSpreadExpression.bid;
                    }
                    else
                    {
                        optionSpreadExpression.defaultBidPriceBeforeTheor = defaultPrice;
                    }


                }
            }
        }

        public void manageExpressionPriceCalcs(OptionSpreadExpression optionSpreadExpression)
        {

            if (optionSpreadExpression.optionExpressionType == OPTION_EXPRESSION_TYPES.OPTION_EXPRESSION_RISK_FREE_RATE)
            {
                // fill all expressions with the latest interest rate

                foreach (KeyValuePair<string, OptionSpreadExpression> entry in optionSpreadExpressionHashTable_keyCQGInId)
                {
                    entry.Value.riskFreeRateFilled = true;

                    entry.Value.riskFreeRate = optionSpreadExpression.riskFreeRate;

                    //only update if subscribed
                    if (entry.Value.setSubscriptionLevel)
                    {
                        entry.Value.totalCalcsRefresh = CQG_REFRESH_STATE.DATA_UPDATED;
                    }
                }
            }
            else
            {
                if (optionSpreadExpression.callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                {

                    optionSpreadExpression.impliedVol = 0;
                    optionSpreadExpression.delta = 1 * DataCollectionConstants.OPTION_DELTA_MULTIPLIER;

                    fillFutureDecisionAndTransactionPrice(optionSpreadExpression);

                    foreach (OptionSpreadExpression optionSpreadThatUsesFuture in optionSpreadExpression.optionExpressionsThatUseThisFutureAsUnderlying)
                    {

                        fillEodAnalysisPrices(optionSpreadThatUsesFuture);

                        fillTheoreticalOptionPrice(optionSpreadThatUsesFuture);

                        generatingGreeks(optionSpreadThatUsesFuture);

                        if (optionSpreadThatUsesFuture.setSubscriptionLevel)
                        {
                            optionSpreadThatUsesFuture.totalCalcsRefresh = CQG_REFRESH_STATE.DATA_UPDATED;
                        }

                        //tempListOfExpressionsFilled.Add(optionSpreadExpression.optionExpressionIdxUsedInFuture[expressionCounter]);
                    }


                }
                else
                {
                    generatingGreeks(optionSpreadExpression);
                }



                optionSpreadExpression.totalCalcsRefresh = CQG_REFRESH_STATE.DATA_UPDATED;
            }
        }

        public void fillFutureDecisionAndTransactionPrice(OptionSpreadExpression optionSpreadExpression)
        {
            if (optionSpreadExpression.decisionBar != null && optionSpreadExpression.todayTransactionBar != null)
            {

                optionSpreadExpression.decisionPrice =
                    optionSpreadExpression.decisionBar.close;

                optionSpreadExpression.decisionPriceTime =
                    optionSpreadExpression.decisionBar.barTime;

                optionSpreadExpression.decisionPriceFilled = true;

                //if (optionSpreadExpression.reachedDecisionBar)
                //{
                //    optionSpreadExpression.filledAfterReachedDecisionBar = true;
                //}




                optionSpreadExpression.transactionPrice =
                    optionSpreadExpression.todayTransactionBar.close;

                optionSpreadExpression.transactionPriceTime =
                    optionSpreadExpression.todayTransactionBar.barTime;


                if (optionSpreadExpression.reachedTransactionTimeBoundary)
                {
                    //optionSpreadExpression.filledAfterTransactionTimeBoundary = true;

                    optionSpreadExpression.transactionPriceFilled = true;

                    foreach (OptionSpreadExpression ose in optionSpreadExpression.optionExpressionsThatUseThisFutureAsUnderlying)
                    {
                        ose.transactionPriceFilled = true;
                    }

                }




            }

        }

        public void fillEodAnalysisPrices(OptionSpreadExpression optionSpreadExpression)
        {

            generatingSettlementGreeks(optionSpreadExpression);

            if (optionSpreadExpression.callPutOrFuture != OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
            {
                //optionSpreadExpression.strikeLevel =
                //        optionSpreadManager.strikeLevelCalc(
                //                optionSpreadExpression.underlyingTransactionTimePrice,
                //                optionSpreadExpression.strikePrice,
                //                optionSpreadExpression.instrument);
                OptionSpreadExpression futureExpression = optionSpreadExpression.underlyingFutureExpression;

                //optionSpreadExpression.theoreticalOptionPrice =

                //BELOW IS CHANGE FOR PREVIOUS SETTLEMENT
                //OCT 23 2014

                optionSpreadExpression.decisionPrice =
                    OptionCalcs.blackScholes(
                        optionSpreadExpression.callPutOrFutureChar,
                           futureExpression.decisionPrice,
                           optionSpreadExpression.strikePrice,
                           optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                           optionSpreadExpression.impliedVolFromSpan);

                double optionTickSize = DataManagementUtility.chooseOptionTickSize(
                            optionSpreadExpression.decisionPrice,
                            optionSpreadExpression.instrument.optionticksize,
                            optionSpreadExpression.instrument.secondaryoptionticksize,
                            optionSpreadExpression.instrument.secondaryoptionticksizerule);

                optionSpreadExpression.decisionPrice =
                    ((int)((optionSpreadExpression.decisionPrice + optionTickSize / 2) /
                        optionTickSize)) * optionTickSize;

                //StringBuilder testingOut = new StringBuilder();
                //testingOut.Append("DECISION,");
                //testingOut.Append(optionSpreadExpression.instrument.name);
                //testingOut.Append(",CALL OR PUT, ");
                //testingOut.Append(optionSpreadExpression.callPutOrFutureChar);
                //testingOut.Append(",FUT PRICE, ");
                //testingOut.Append(futureExpression.decisionPrice);
                //testingOut.Append(",STRIKE PRICE, ");
                //testingOut.Append(optionSpreadExpression.strikePrice);
                //testingOut.Append(",YEAR FRACTION, ");
                //testingOut.Append(optionSpreadExpression.yearFraction);
                //testingOut.Append(",RFR, ");
                //testingOut.Append(optionSpreadExpression.riskFreeRate);
                //testingOut.Append(",SETTLE, ");
                //testingOut.Append(optionSpreadExpression.settlementImpliedVol);
                //testingOut.Append(",OPTION PRICE, ");
                //testingOut.Append(optionSpreadExpression.decisionPrice);

                //TSErrorCatch.debugWriteOut(testingOut.ToString());

                optionSpreadExpression.decisionPriceTime = futureExpression.decisionPriceTime;

                optionSpreadExpression.decisionPriceFilled = true;

                optionSpreadExpression.reachedDecisionBar = futureExpression.reachedDecisionBar;

                optionSpreadExpression.reachedBarAfterDecisionBar = futureExpression.reachedBarAfterDecisionBar;

                optionSpreadExpression.reached1MinAfterDecisionBarUsedForSnapshot = futureExpression.reached1MinAfterDecisionBarUsedForSnapshot;

                //optionSpreadExpression.instrument.eodAnalysisAtInstrument = futureExpression.instrument.eodAnalysisAtInstrument;

                //if (optionSpreadManager.realtimeMonitorSettings.eodAnalysis)
                if (optionSpreadExpression.instrument.eodAnalysisAtInstrument)
                {
                    optionSpreadExpression.transactionPrice =
                        OptionCalcs.blackScholes(
                            optionSpreadExpression.callPutOrFutureChar,
                               futureExpression.transactionPrice,
                               optionSpreadExpression.strikePrice,
                               optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                               optionSpreadExpression.settlementImpliedVol);
                }
                else
                {
                    if (!optionSpreadExpression.filledAfterTransactionTimeBoundary)
                    {
                        if (optionSpreadExpression.impliedVolFilled && optionSpreadExpression.transactionPriceFilled)
                        {
                            optionSpreadExpression.transactionPrice =
                                OptionCalcs.blackScholes(
                                    optionSpreadExpression.callPutOrFutureChar,
                                       futureExpression.transactionPrice,
                                       optionSpreadExpression.strikePrice,
                                       optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                                       optionSpreadExpression.impliedVol);

                            optionSpreadExpression.filledAfterTransactionTimeBoundary = true;

                        }
                    }
                }

                optionTickSize = DataManagementUtility.chooseOptionTickSize(
                            optionSpreadExpression.transactionPrice,
                            optionSpreadExpression.instrument.optionticksize,
                            optionSpreadExpression.instrument.secondaryoptionticksize,
                            optionSpreadExpression.instrument.secondaryoptionticksizerule);

                optionSpreadExpression.transactionPrice =
                    ((int)((optionSpreadExpression.transactionPrice + optionTickSize / 2) /
                        optionTickSize)) * optionTickSize;

                //testingOut.Clear();
                //testingOut.Append("TRANS,");
                //testingOut.Append(optionSpreadExpression.instrument.name);
                //testingOut.Append(",CALL OR PUT, ");
                //testingOut.Append(optionSpreadExpression.callPutOrFutureChar);
                //testingOut.Append(",FUT PRICE, ");
                //testingOut.Append(futureExpression.transactionPrice);
                //testingOut.Append(",STRIKE PRICE, ");
                //testingOut.Append(optionSpreadExpression.strikePrice);
                //testingOut.Append(",YEAR FRACTION, ");
                //testingOut.Append(optionSpreadExpression.yearFraction);
                //testingOut.Append(",RFR, ");
                //testingOut.Append(optionSpreadExpression.riskFreeRate);
                //testingOut.Append(",SETTLE, ");
                //testingOut.Append(optionSpreadExpression.settlementImpliedVol);
                //testingOut.Append(",OPTION PRICE, ");
                //testingOut.Append(optionSpreadExpression.transactionPrice);

                //TSErrorCatch.debugWriteOut(testingOut.ToString());

                optionSpreadExpression.transactionPriceTime = futureExpression.transactionPriceTime;

                optionSpreadExpression.decisionPriceFilled = true;

                optionSpreadExpression.reachedTransactionTimeBoundary = futureExpression.reachedTransactionTimeBoundary;

                //optionSpreadExpression.defaultPrice = optionSpreadExpression.theoreticalOptionPrice;

                //CHANGED DEC 30 2015 
                //if (optionSpreadExpression.instrument.eodAnalysisAtInstrument)
                //{
                //    optionSpreadExpression.defaultPrice =
                //        optionSpreadExpression.transactionPrice;

                //    optionSpreadExpression.minutesSinceLastUpdate = 0;

                //    optionSpreadExpression.lastTimeUpdated =
                //        optionSpreadExpression.transactionPriceTime;

                //    optionSpreadExpression.defaultPriceFilled = true;
                //}
            }
        }

        public void fillTheoreticalOptionPrice(OptionSpreadExpression optionSpreadExpression)
        {



            if (optionSpreadExpression.callPutOrFuture != OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
            {
                OptionSpreadExpression futureExpression = optionSpreadExpression.underlyingFutureExpression;

                if (optionSpreadExpression.theoreticalOptionDataList == null)
                {
                    optionSpreadExpression.theoreticalOptionDataList = new List<TheoreticalBar>();
                }

                if (optionSpreadExpression.riskFreeRateFilled
                    && futureExpression.futureBarData != null
                    && (futureExpression.futureBarData.Count - 1) > optionSpreadExpression.theoreticalOptionDataList.Count)
                {
                    for (int i = optionSpreadExpression.theoreticalOptionDataList.Count; i < futureExpression.futureBarData.Count - 1;
                        i++)
                    {
                        TheoreticalBar bar = new TheoreticalBar();

                        bar.barTime = futureExpression.futureBarData[i].barTime;

                        bar.price =
                            OptionCalcs.blackScholes(
                            optionSpreadExpression.callPutOrFutureChar,
                               futureExpression.futureBarData[i].close,
                               optionSpreadExpression.strikePrice,
                               optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                               optionSpreadExpression.impliedVolFromSpan);  // / 100);

                        double optionTickSize = DataManagementUtility.chooseOptionTickSize(
                            bar.price,
                            optionSpreadExpression.instrument.optionticksize,
                            optionSpreadExpression.instrument.secondaryoptionticksize,
                            optionSpreadExpression.instrument.secondaryoptionticksizerule);

                        bar.price =
                            ((int)((bar.price + optionTickSize / 2) /
                                optionTickSize)) * optionTickSize;

                        optionSpreadExpression.theoreticalOptionDataList.Add(bar);

                        //if (futureExpression.futureBarData[i].close != 0)

                        //TSErrorCatch.debugWriteOut(futureExpression.futureBarData[i].close + "," + bar.price);
                    }
                }

                //if (optionSpreadExpression.cqgSymbol.CompareTo("C.US.TYAM1412600") == 0)
                //{
                //    TSErrorCatch.debugWriteOut("TEST");
                //}

                if (futureExpression.defaultPriceFilled)
                //CHANGED DEC 30 2015 
                //&& !optionSpreadExpression.instrument.eodAnalysisAtInstrument)
                {


                    {
                        //optionSpreadExpression.strikeLevel =
                        //        optionSpreadManager.strikeLevelCalc(
                        //                futureExpression.defaultPrice,
                        //                optionSpreadExpression.strikePrice,
                        //                optionSpreadExpression.instrument);

                        optionSpreadExpression.theoreticalOptionPrice =
                            OptionCalcs.blackScholes(
                            optionSpreadExpression.callPutOrFutureChar,
                               futureExpression.defaultPrice,
                               optionSpreadExpression.strikePrice,
                               optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                               optionSpreadExpression.impliedVolFromSpan);  // / 100);

                        double optionTickSize = DataManagementUtility.chooseOptionTickSize(
                            optionSpreadExpression.theoreticalOptionPrice,
                            optionSpreadExpression.instrument.optionticksize,
                            optionSpreadExpression.instrument.secondaryoptionticksize,
                            optionSpreadExpression.instrument.secondaryoptionticksizerule);

                        optionSpreadExpression.theoreticalOptionPrice =
                            ((int)((optionSpreadExpression.theoreticalOptionPrice + optionTickSize / 2) /
                                optionTickSize)) * optionTickSize;

                        if (optionSpreadExpression.theoreticalOptionPrice == 0)
                        {
                            optionSpreadExpression.theoreticalOptionPrice = optionSpreadExpression.instrument.optionticksize; // OptionConstants.OPTION_ZERO_PRICE;
                        }

                        if (realtimePriceFillType
                                == REALTIME_PRICE_FILL_TYPE.PRICE_MID_BID_ASK)
                        {
                            optionSpreadExpression.defaultPrice = optionSpreadExpression.defaultMidPriceBeforeTheor;
                        }
                        else if (realtimePriceFillType
                                == REALTIME_PRICE_FILL_TYPE.PRICE_ASK)
                        {
                            optionSpreadExpression.defaultPrice = optionSpreadExpression.defaultAskPriceBeforeTheor;
                        }
                        else if (realtimePriceFillType
                                == REALTIME_PRICE_FILL_TYPE.PRICE_BID)
                        {
                            optionSpreadExpression.defaultPrice = optionSpreadExpression.defaultBidPriceBeforeTheor;
                        }
                        else if (realtimePriceFillType
                                == REALTIME_PRICE_FILL_TYPE.PRICE_DEFAULT)
                        {
                            bool midPriceAcceptable = false;

                            if (optionSpreadExpression.bidFilled && optionSpreadExpression.askFilled
                                &&
                                Math.Abs((optionSpreadExpression.bid - optionSpreadExpression.ask)
                            / DataManagementUtility.chooseOptionTickSize(optionSpreadExpression.ask,
                                optionSpreadExpression.instrument.optionticksize,
                                optionSpreadExpression.instrument.secondaryoptionticksize,
                                optionSpreadExpression.instrument.secondaryoptionticksizerule))
                            < DataCollectionConstants.OPTION_ACCEPTABLE_BID_ASK_SPREAD)
                            {
                                midPriceAcceptable = true;
                            }

                            if (midPriceAcceptable
                                ||
                                Math.Abs((optionSpreadExpression.defaultMidPriceBeforeTheor
                                - optionSpreadExpression.theoreticalOptionPrice)
                                / optionSpreadExpression.theoreticalOptionPrice) <= DataCollectionConstants.OPTION_DEFAULT_THEORETICAL_PRICE_RANGE)
                            {
                                //                         TSErrorCatch.debugWriteOut(optionSpreadExpression.cqgSymbol + "  NOT theoretical Price "
                                //                             + optionSpreadExpression.optionDefaultPriceWithoutTheoretical);

                                //optionSpreadExpression.underlyingDefaultPrice = defaultPrice;
                                //if (optionSpreadExpression.callPutOrFuture != (int)OPTION_SPREAD_CONTRACT_TYPE.CALL)

                                bool filledDefaultPrice = false;

                                if (optionSpreadExpression.callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.CALL)
                                {
                                    if (futureExpression.defaultPrice >
                                        optionSpreadExpression.strikePrice + (DataCollectionConstants.STRIKE_COUNT_FOR_DEFAULT_TO_THEORETICAL
                                            * optionSpreadExpression.instrument.optionstrikeincrement))
                                    {
                                        optionSpreadExpression.defaultPrice = optionSpreadExpression.theoreticalOptionPrice;
                                        filledDefaultPrice = true;
                                    }
                                }
                                else if (optionSpreadExpression.callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.PUT)
                                {
                                    if (futureExpression.defaultPrice <
                                        optionSpreadExpression.strikePrice - (DataCollectionConstants.STRIKE_COUNT_FOR_DEFAULT_TO_THEORETICAL
                                            * optionSpreadExpression.instrument.optionstrikeincrement))
                                    {
                                        optionSpreadExpression.defaultPrice = optionSpreadExpression.theoreticalOptionPrice;
                                        filledDefaultPrice = true;
                                    }
                                }

                                if (!filledDefaultPrice)
                                {
                                    optionSpreadExpression.defaultPrice = optionSpreadExpression.defaultMidPriceBeforeTheor;
                                }
                            }
                            else
                            {
                                /*TSErrorCatch.debugWriteOut(optionSpreadExpression.cqgSymbol + "  IS theoretical Price");*/

                                optionSpreadExpression.defaultPrice = optionSpreadExpression.theoreticalOptionPrice;
                            }
                        }
                        else
                        {
                            optionSpreadExpression.defaultPrice = optionSpreadExpression.theoreticalOptionPrice;
                        }
                    }


                    optionSpreadExpression.defaultPriceFilled = true;

                    //optionSpreadExpression.defaultPriceForDisplay = optionSpreadExpression.defaultPrice;
                }
            }
        }

        private void fillPricesFromQuote(OptionSpreadExpression optionSpreadExpression, CQGQuotes quotes)
        {
            //double defaultPrice = 0;

            //CQGQuote quoteAsk = optionSpreadExpression.cqgInstrument.Quotes[eQuoteType.qtAsk];
            //CQGQuote quoteBid = optionSpreadExpression.cqgInstrument.Quotes[eQuoteType.qtBid];
            //CQGQuote quoteTrade = optionSpreadExpression.cqgInstrument.Quotes[eQuoteType.qtTrade];
            //CQGQuote quoteSettlement = optionSpreadExpression.cqgInstrument.Quotes[eQuoteType.qtSettlement];
            //CQGQuote quoteYestSettlement = optionSpreadExpression.cqgInstrument.Quotes[eQuoteType.qtYesterdaySettlement];

            CQGQuote quoteAsk = quotes[eQuoteType.qtAsk];
            CQGQuote quoteBid = quotes[eQuoteType.qtBid];
            CQGQuote quoteTrade = quotes[eQuoteType.qtTrade];
            CQGQuote quoteSettlement = quotes[eQuoteType.qtSettlement];
            CQGQuote quoteYestSettlement = quotes[eQuoteType.qtYesterdaySettlement];

            if (optionSpreadExpression.callPutOrFuture != OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
            {
                if (quoteAsk != null)
                {
                    if (quoteAsk.IsValid)
                    {
                        optionSpreadExpression.ask = quoteAsk.Price;

                        optionSpreadExpression.askFilled = true;
                    }
                    else
                    {
                        optionSpreadExpression.ask = 0;

                        optionSpreadExpression.askFilled = false;
                    }
                }

                if (quoteBid != null)
                {
                    if (quoteBid.IsValid)
                    {
                        optionSpreadExpression.bid = quoteBid.Price;

                        optionSpreadExpression.bidFilled = true;
                    }
                    else
                    {
                        optionSpreadExpression.bid = 0;

                        optionSpreadExpression.bidFilled = false;
                    }
                }

                if (quoteTrade != null)
                {
                    if (quoteTrade.IsValid)
                    {
                        optionSpreadExpression.trade = quoteTrade.Price;

                        optionSpreadExpression.tradeFilled = true;
                    }
                    else if (optionSpreadExpression.callPutOrFuture != OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
                    {
                        optionSpreadExpression.trade = 0;

                        optionSpreadExpression.tradeFilled = false;
                    }
                }
            }

            if (quoteSettlement != null)
            {
                if (quoteSettlement.IsValid)
                {
                    //if (optionSpreadManager.realtimeMonitorSettings.eodAnalysis
                    //    || (optionSpreadExpression.instrument != null
                    //        && optionSpreadExpression.instrument.eodAnalysisAtInstrument))
                    //{
                    //    if (optionSpreadExpression.substituteSubscriptionRequest
                    //        || !optionSpreadExpression.useSubstituteSymbolAtEOD)
                    //    {
                    //        if (!optionSpreadExpression.manuallyFilled)
                    //        {
                    //            optionSpreadExpression.settlement = quoteSettlement.Price;

                    //            optionSpreadExpression.settlementDateTime = quoteSettlement.ServerTimestamp;
                    //        }

                    //    }
                    //}
                    //else
                    //{
                    //    if (!optionSpreadExpression.manuallyFilled)
                    //    {
                    //        optionSpreadExpression.settlement = quoteSettlement.Price;

                    //        optionSpreadExpression.settlementDateTime = quoteSettlement.ServerTimestamp;
                    //    }
                    //}

                    optionSpreadExpression.settlement = quoteSettlement.Price;

                    optionSpreadExpression.settlementDateTime = quoteSettlement.ServerTimestamp;

                    if (optionSpreadExpression.settlementDateTime.Date.CompareTo(DateTime.Now.Date) == 0)
                    {
                        optionSpreadExpression.settlementIsCurrentDay = true;
                    }


                    optionSpreadExpression.settlementFilled = true;

                    //fillEODSubstitutePrices(optionSpreadExpression);

                    //if (optionSpreadManager.realtimeMonitorSettings.eodAnalysis
                    //    && optionSpreadExpression.substituteSubscriptionRequest)
                    //{
                    //    foreach(OptionSpreadExpression optionSpreadExpressionSingle in 
                    //        optionSpreadExpression.mainExpressionSubstitutionUsedFor)
                    //    {
                    //        optionSpreadExpressionSingle.settlement = optionSpreadExpression.settlement;

                    //        optionSpreadExpressionSingle.settlementDateTime = optionSpreadExpression.settlementDateTime;

                    //        optionSpreadExpressionSingle.settlementIsCurrentDay = optionSpreadExpression.settlementIsCurrentDay;

                    //        optionSpreadExpressionSingle.settlementFilled = true;
                    //    }
                    //}

                    //TSErrorCatch.debugWriteOut(quoteSettlement.Timestamp + "  " + quoteSettlement.ServerTimestamp);
                }
                else
                {
                    //if (!optionSpreadExpression.manuallyFilled)
                    {
                        optionSpreadExpression.settlement = 0;

                        optionSpreadExpression.settlementFilled = false;
                    }
                }


            }

            //if (quoteYestSettlement != null)
            //{
            //    if (quoteYestSettlement.IsValid)
            //    {
            //        optionSpreadExpression.yesterdaySettlement = quoteYestSettlement.Price;

            //        optionSpreadExpression.yesterdaySettlementFilled = true;
            //    }
            //    else
            //    {
            //        optionSpreadExpression.yesterdaySettlement = 0;

            //        optionSpreadExpression.yesterdaySettlementFilled = false;
            //    }


            //}





            //if (optionSpreadExpression.callPutOrFuture
            //                        != (int)OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
            //{
            //    TSErrorCatch.debugWriteOut("fillPricesFromQuote "
            //        + optionSpreadExpression.cqgSymbol
            //        + " bid " + optionSpreadExpression.bid
            //        + " ask " + optionSpreadExpression.ask
            //        + " trade " + optionSpreadExpression.trade);
            //}

        }


        public void generatingGreeks(OptionSpreadExpression optionSpreadExpression)
        {

            if (optionSpreadExpression.callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
            {
                optionSpreadExpression.impliedVol = 0;
                optionSpreadExpression.delta = 1 * DataCollectionConstants.OPTION_DELTA_MULTIPLIER;
            }

            else
            {

                OptionSpreadExpression futureExpression = optionSpreadExpression.underlyingFutureExpression;

                if (optionSpreadExpression.defaultPriceFilled
                    && futureExpression.defaultPriceFilled)
                {
                    //                 optionSpreadExpression.settlementImpliedVol =
                    //                     OptionCalcs.calculateOptionVolatilityNR(
                    //                        optionSpreadExpression.callPutOrFutureChar,
                    //                        optionSpreadExpression.underlyingSettlementPrice,
                    //                        optionSpreadExpression.strikePrice,
                    //                        optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                    //                        optionSpreadExpression.settlement) * 100;



                    //if (!optionSpreadManager.realtimeMonitorSettings.eodAnalysis)

                    double optionTickSize = optionSpreadExpression.instrument.optionticksize;

                    if (optionSpreadExpression.instrument.secondaryoptionticksize > 0)
                    {
                        optionTickSize = optionSpreadExpression.instrument.secondaryoptionticksize;
                    }

                    optionSpreadExpression.impliedVol =
                        OptionCalcs.calculateOptionVolatilityNR(
                           optionSpreadExpression.callPutOrFutureChar,
                           futureExpression.defaultPrice,
                           optionSpreadExpression.strikePrice,
                           optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                           optionSpreadExpression.defaultPrice, optionTickSize); // *100;


                    if (optionSpreadExpression.impliedVol > 1000
                        || optionSpreadExpression.impliedVol < -1000
                        )
                    {
                        optionSpreadExpression.defaultPrice = optionSpreadExpression.theoreticalOptionPrice;

                        optionSpreadExpression.impliedVol = optionSpreadExpression.impliedVolFromSpan;

                        //optionSpreadExpression.impliedVol = optionSpreadExpression.impliedVolIndexShifted;



                        //                     OptionCalcs.calculateOptionVolatilityNR(
                        //                        optionSpreadExpression.callPutOrFutureChar,
                        //                        optionSpreadExpression.underlyingDefaultPrice,
                        //                        optionSpreadExpression.strikePrice,
                        //                        optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                        //                        optionSpreadExpression.optionDefaultPrice);

                        optionSpreadExpression.defaultPriceFilled = true;

                        //optionSpreadExpression.defaultPriceForDisplay = optionSpreadExpression.defaultPrice;

                        optionSpreadExpression.totalCalcsRefresh = CQG_REFRESH_STATE.DATA_UPDATED;
                    }

                    optionSpreadExpression.impliedVolFilled = true;

                    optionSpreadExpression.delta =
                        OptionCalcs.gDelta(
                           optionSpreadExpression.callPutOrFutureChar,
                           futureExpression.defaultPrice,
                           optionSpreadExpression.strikePrice,
                           optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                           0,
                           optionSpreadExpression.impliedVol) * DataCollectionConstants.OPTION_DELTA_MULTIPLIER;

                    //optionSpreadExpression.gamma =
                    //    OptionCalcs.gGamma(
                    //    //optionSpreadExpression.callPutOrFutureChar,
                    //       futureExpression.defaultPrice,
                    //       optionSpreadExpression.strikePrice,
                    //       optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                    //       0,
                    //       optionSpreadExpression.impliedVol);

                    //optionSpreadExpression.vega =
                    //    OptionCalcs.gVega(
                    //    //optionSpreadExpression.callPutOrFutureChar,
                    //       futureExpression.defaultPrice,
                    //       optionSpreadExpression.strikePrice,
                    //       optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                    //       0,
                    //       optionSpreadExpression.impliedVol);

                    //optionSpreadExpression.theta =
                    //    OptionCalcs.gTheta(
                    //        optionSpreadExpression.callPutOrFutureChar,
                    //       futureExpression.defaultPrice,
                    //       optionSpreadExpression.strikePrice,
                    //       optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                    //       0,
                    //       optionSpreadExpression.impliedVol);
                }
            }

        }

        public void generatingSettlementGreeks(OptionSpreadExpression optionSpreadExpression)
        //OptionSpreadExpression futureExpression)
        {



            //if (futureExpression.settlementFilled
            if (optionSpreadExpression.settlementFilled)
            {

                OptionSpreadExpression futureExpression = optionSpreadExpression.underlyingFutureExpression;

                if (optionSpreadExpression.settlementFilled)
                {

                    double optionTickSize = optionSpreadExpression.instrument.optionticksize;

                    if (optionSpreadExpression.instrument.secondaryoptionticksize > 0)
                    {
                        optionTickSize = optionSpreadExpression.instrument.secondaryoptionticksize;
                    }

                    optionSpreadExpression.settlementImpliedVol =
                        OptionCalcs.calculateOptionVolatilityNR(
                           optionSpreadExpression.callPutOrFutureChar,
                           futureExpression.settlement,
                           optionSpreadExpression.strikePrice,
                           optionSpreadExpression.yearFraction, optionSpreadExpression.riskFreeRate,
                           optionSpreadExpression.settlement, optionTickSize);// *100;
                }

                //if (optionBuildCommonMethods.usingEODSettlements)
                //{
                //optionSpreadThatUsesFuture.impliedVol = optionSpreadExpression.settlementImpliedVol;
                //}


            }

        }
    }
}
