using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MongoDB.Bson;
using MongoDB.Driver;
using System.Net;
using System.Net.Sockets;

namespace DataCollectionForRealtime
{
    public partial class RealtimeDataManagement : Form
    {
        private CQGDataManagement cqgDataManagement;

        public RealtimeDataManagement()
        {
            InitializeComponent();

            cqgDataManagement = new CQGDataManagement(this);

            MongoDBConnectionAndSetup mongoDBConnectionAndSetup = new MongoDBConnectionAndSetup();
            //mongoDBConnectionAndSetup.connectToMongoDB();
            //mongoDBConnectionAndSetup.createDocument();
            //mongoDBConnectionAndSetup.dropCollection();

            var contextTMLDB = new DataClassesTMLDBDataContext(
                System.Configuration.ConfigurationManager.ConnectionStrings["TMLDBConnectionString"].ConnectionString);
            TMLDBReader TMLDBReader = new TMLDBReader(contextTMLDB);

            

            bool gotInstrumentList = TMLDBReader.GetTblInstruments(ref cqgDataManagement.instrumentHashTable,
                    ref cqgDataManagement.instrumentList);

            Console.WriteLine(cqgDataManagement.instrumentList[0].description);

            AsyncTaskListener.Updated += AsyncTaskListener_Updated;

            testLoadIn();

            testGetData();

            //mongoDBConnectionAndSetup.createDoc();
            //mongoDBConnectionAndSetup.getDocument();
        }

        private void testLoadIn()
        {
            MongoDBConnectionAndSetup mongoDBConnectionAndSetup = new MongoDBConnectionAndSetup();

            Mongo_OptionSpreadExpression osefdb = new Mongo_OptionSpreadExpression();

            
            osefdb.cqgSymbol = "F.EPU16";
            osefdb.instrument = cqgDataManagement.instrumentHashTable[11];

            //mongoDBConnectionAndSetup.MongoDataCollection.ReplaceOne(
            //    item => item.cqgSymbol == osefdb.cqgSymbol,
            //    osefdb,
            //    new UpdateOptions { IsUpsert = true });

            mongoDBConnectionAndSetup.MongoDataCollection.InsertOne(osefdb);


        }

        private void testGetData()
        {
            MongoDBConnectionAndSetup mongoDBConnectionAndSetup = new MongoDBConnectionAndSetup();

            var filterBuilder = Builders<Mongo_OptionSpreadExpression>.Filter;
            var filter = filterBuilder.Ne("Id", "barf");

            var testExpression = mongoDBConnectionAndSetup.MongoDataCollection.Find(filter).SingleOrDefault();

            Console.WriteLine(testExpression.cqgSymbol);


        }

        private void btnCallAllInstruments_Click(object sender, EventArgs e)
        {

        }

        internal void updateConnectionStatus(string connectionStatusLabel,
            Color connColor)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke((MethodInvoker)delegate()
                { 
                    connectionStatus.Text = connectionStatusLabel;
                    connectionStatus.ForeColor = connColor;
                });
            }
            else
            {
                connectionStatus.Text = connectionStatusLabel;
                connectionStatus.ForeColor = connColor;
            }
        }

        internal void updateCQGDataStatus(String dataStatus, Color backColor, Color foreColor)
        {
#if DEBUG
            try
#endif
            {
                if (this.InvokeRequired)
                {
                    this.BeginInvoke((MethodInvoker)delegate()
                    {
                        this.dataStatus.ForeColor = foreColor;
                        this.dataStatus.BackColor = backColor;
                        this.dataStatus.Text = dataStatus;
                    });
                }
                else
                {
                    this.dataStatus.ForeColor = foreColor;
                    this.dataStatus.BackColor = backColor;
                    this.dataStatus.Text = dataStatus;
                }
            }
#if DEBUG
            catch (Exception ex)
            {
                TSErrorCatch.errorCatchOut(Convert.ToString(this), ex);
            }
#endif
        }

        public void updateStatusSubscribeData(String subcriptionMessage)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke((MethodInvoker)delegate()
                {
                    statusSubscribeData.Text = subcriptionMessage;
                });
            }
            else
            {
                statusSubscribeData.Text = subcriptionMessage;
            }
        }

        /// <summary>
        /// Set up the realtime price fill type
        /// </summary>
        private void optionPriceFillTypeChanged()
        {
            if(radioBtnDefaultPriceRules.Checked)
            {
                cqgDataManagement.realtimePriceFillType = REALTIME_PRICE_FILL_TYPE.PRICE_DEFAULT;
            }
            else if (radioBtnMidPriceRules.Checked)
            {
                cqgDataManagement.realtimePriceFillType = REALTIME_PRICE_FILL_TYPE.PRICE_MID_BID_ASK;
            }
            else if (radioBtnTheorPriceRules.Checked)
            {
                cqgDataManagement.realtimePriceFillType = REALTIME_PRICE_FILL_TYPE.PRICE_THEORETICAL;
            }
            else if (radioBtnAskPriceRules.Checked)
            {
                cqgDataManagement.realtimePriceFillType = REALTIME_PRICE_FILL_TYPE.PRICE_ASK;
            }
            else if (radioBtnBidPriceRules.Checked)
            {
                cqgDataManagement.realtimePriceFillType = REALTIME_PRICE_FILL_TYPE.PRICE_BID;
            }

            //realtimeMonitorSettings.realtimePriceFillType = realtimePriceFillType;

            //optionSpreadManager.updatedRealtimeMonitorSettingsThreadRun();

            //optionRealtimeMonitor.updateStatusStripOptionMonitor();
        }

        private void radioBtnDefaultPriceRules_CheckedChanged(object sender, EventArgs e)
        {
            optionPriceFillTypeChanged();
        }
        
        private void AsyncTaskListener_Updated(
    string message = null,
    int progress = -1,
    double rps = double.NaN)
        {
            Action action = new Action(
                () =>
                {
                    if (!string.IsNullOrWhiteSpace(message))
                    {
                        richTextBoxLog.Text += message + "\n";
                        richTextBoxLog.Select(richTextBoxLog.Text.Length, richTextBoxLog.Text.Length);
                        richTextBoxLog.ScrollToCaret();
                    }
                    //if (progress != -1)
                    //{
                    //    progressBar.Value = progress;
                    //}
                    //if (!double.IsNaN(rps))
                    //{
                    //    labelRPS2.Text = Math.Round(rps).ToString();
                    //}
                });

            try
            {
                Invoke(action);
            }
            catch (ObjectDisposedException)
            {
                // User closed the form
            }
        }

    }
}
