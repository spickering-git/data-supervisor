using System;
using System.Drawing;
using System.Text;
using System.Threading;
using CQG;

namespace DataCollectionForRealtime
{
    class CQGDataManagement
    {

        internal CQGDataManagement(RealtimeDataManagement realtimeDataManagement)
        {
            this.realtimeDataManagement = realtimeDataManagement;

            ThreadPool.QueueUserWorkItem(new WaitCallback(initializeCQGAndCallbacks));
        }

        //private DataManagementUtility dataManagementUtility = new DataManagementUtility();

        //private Thread subscriptionThread;
        //private bool subscriptionThreadShouldStop = false;
        private const int SUBSCRIPTION_TIMEDELAY_CONSTANT = 125;

        private RealtimeDataManagement realtimeDataManagement;
        
        private CQG.CQGCEL m_CEL;

        public CQG.CQGCEL M_CEL
        {
            get { return m_CEL; }
            set { M_CEL = value; }
        }

        //List<OptionSpreadExpression> optionSpreadExpressionList = new List<OptionSpreadExpression>();

        //private ConcurrentDictionary<string, OptionSpreadExpression> optionSpreadExpressionHashTable_keySymbol
        //    = new ConcurrentDictionary<string, OptionSpreadExpression>();

        //private ConcurrentDictionary<string, OptionSpreadExpression> optionSpreadExpressionHashTable_keyCQGInId
        //    = new ConcurrentDictionary<string, OptionSpreadExpression>();

        //private ConcurrentDictionary<string, OptionSpreadExpression> optionSpreadExpressionHashTable_keyFullName
        //    = new ConcurrentDictionary<string, OptionSpreadExpression>();

        //internal ConcurrentDictionary<long, Instrument> instrumentHashTable
        //    = new ConcurrentDictionary<long, Instrument>();

        //internal List<Instrument> instrumentList = new List<Instrument>();


        internal void connectCQG()
        {
            if (m_CEL != null)
            {
                m_CEL.Startup();
                Type t = m_CEL.GetType();
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

                //m_CEL.TimedBarsResolved += new CQG._ICQGCELEvents_TimedBarsResolvedEventHandler(m_CEL_TimedBarResolved);
                //m_CEL.TimedBarsAdded += new CQG._ICQGCELEvents_TimedBarsAddedEventHandler(m_CEL_TimedBarsAdded);
                //m_CEL.TimedBarsUpdated += new CQG._ICQGCELEvents_TimedBarsUpdatedEventHandler(m_CEL_TimedBarsUpdated);

                ////m_CEL.IncorrectSymbol += new _ICQGCELEvents_IncorrectSymbolEventHandler(CEL_IncorrectSymbol);
                //m_CEL.InstrumentSubscribed += new _ICQGCELEvents_InstrumentSubscribedEventHandler(m_CEL_InstrumentSubscribed);
                //m_CEL.InstrumentChanged += new _ICQGCELEvents_InstrumentChangedEventHandler(m_CEL_InstrumentChanged);

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

        private bool isCQGStarted;

        public bool IsCQGStarted
        {
            get
            {
                isCQGStarted = m_CEL.IsStarted;
                return isCQGStarted;
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

//        public void sendSubscribeRequest(bool sendOnlyUnsubscribed)
//        {

//#if DEBUG
//            try
//#endif
//            {
//                subscriptionThread = new Thread(new ParameterizedThreadStart(sendSubscribeRequestRun));
//                subscriptionThread.IsBackground = true;
//                subscriptionThread.Start(sendOnlyUnsubscribed);

//            }
//#if DEBUG
//            catch (Exception ex)
//            {
//                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
//            }
//#endif

//        }

        //public void sendSubscribeRequestRun(Object obj)
        //{
        //    dataManagementUtility.openThread(null, null);

        //    try
        //    {
        //        //m_CEL.RemoveAllTimedBars();
        //        //Thread.Sleep(3000);

        //        if (m_CEL.IsStarted)
        //        {
        //            bool sendOnlyUnsubscribed = (bool)obj;

        //            int i = 0;

        //            //m_CEL.NewInstrument("C.US.EPN1312600");

        //            while (!subscriptionThreadShouldStop && i < optionSpreadExpressionList.Count)
        //            {

        //                //TSErrorCatch.debugWriteOut("SUBSCRIBE " + optionSpreadExpressionList[i].cqgSymbol);

        //                if (sendOnlyUnsubscribed)
        //                {


        //                    if (!optionSpreadExpressionList[i].setSubscriptionLevel)
        //                    {
        //                        Thread.Sleep(SUBSCRIPTION_TIMEDELAY_CONSTANT);

        //                        realtimeDataManagement.updateStatusSubscribeData(
        //                            "SUBSCRIBE " + optionSpreadExpressionList[i].cqgSymbol
        //                            + " : " + i + " OF " +
        //                            optionSpreadExpressionList.Count);

        //                        m_CEL.NewInstrument(optionSpreadExpressionList[i].cqgSymbol);

        //                    }

        //                    if (optionSpreadExpressionList[i].callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
        //                    {
        //                        if (!optionSpreadExpressionList[i].requestedMinuteBars
        //                            && optionSpreadExpressionList[i].normalSubscriptionRequest)
        //                        {
        //                            requestFutureContractTimeBars(optionSpreadExpressionList[i]);
        //                        }
        //                    }

        //                }
        //                else
        //                {
        //                    optionSpreadExpressionList[i].setSubscriptionLevel = false;

        //                    Thread.Sleep(SUBSCRIPTION_TIMEDELAY_CONSTANT);

        //                    realtimeDataManagement.updateStatusSubscribeData(
        //                            "SUBSCRIBE " + optionSpreadExpressionList[i].cqgSymbol
        //                            + " : " + i + " OF " +
        //                            optionSpreadExpressionList.Count);

        //                    m_CEL.NewInstrument(optionSpreadExpressionList[i].cqgSymbol);

        //                    if (optionSpreadExpressionList[i].callPutOrFuture == OPTION_SPREAD_CONTRACT_TYPE.FUTURE)
        //                    {
        //                        optionSpreadExpressionList[i].requestedMinuteBars = false;                                

        //                        if (optionSpreadExpressionList[i].normalSubscriptionRequest)
        //                        {
        //                            requestFutureContractTimeBars(optionSpreadExpressionList[i]);
        //                        }

        //                    }
        //                }



        //                i++;
        //            }

        //            Thread.Sleep(SUBSCRIPTION_TIMEDELAY_CONSTANT);

        //            realtimeDataManagement.updateStatusSubscribeData("");

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
        //    }

        //    dataManagementUtility.closeThread(null, null);
        //}

        //public void requestFutureContractTimeBars(OptionSpreadExpression optionSpreadExpression)
        //{
        //    try
        //    {

        //        CQGTimedBarsRequest timedBarsRequest = m_CEL.CreateTimedBarsRequest();

        //        timedBarsRequest.Symbol = optionSpreadExpression.cqgSymbol;

        //        timedBarsRequest.SessionsFilter = 31;

        //        timedBarsRequest.IntradayPeriod = 1;

        //        timedBarsRequest.Continuation = CQG.eTimeSeriesContinuationType.tsctNoContinuation;
        //        //do not want continuation bars

        //        DateTime rangeStart = optionSpreadExpression.previousDateTimeBoundaryStart;

        //        DateTime rangeEnd = m_CEL.Environment.LineTime;

        //        timedBarsRequest.RangeStart = rangeStart;

        //        timedBarsRequest.RangeEnd = rangeEnd;

        //        timedBarsRequest.IncludeEnd = true;

        //        timedBarsRequest.UpdatesEnabled = true;

        //        optionSpreadExpression.futureTimedBars = m_CEL.RequestTimedBars(timedBarsRequest);


        //        optionSpreadExpressionHashTable_keyCQGInId.AddOrUpdate(
        //            optionSpreadExpression.futureTimedBars.Id,
        //            optionSpreadExpression, (oldKey, oldValue) => optionSpreadExpression);
        //    }
        //    catch (Exception ex)
        //    {
        //        TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
        //    }
        //}

        

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

    }
}
