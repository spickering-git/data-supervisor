using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CQG;

namespace UnitTestCQG
{
    [TestClass]
    public class UnitTest1
    {
        Dictionary<string, string> histSession;
        List<Dictionary<string, string>> histSessions;

        // Specifies the not available string
        private const string N_A = "N/A";

        [field: System.CLSCompliant(false)]
        public CQGCEL m_CEL;

        private void frmHistoricalSessions_Load(System.Object sender, System.EventArgs e)
        {

            // Creates the CQGCEL object
            m_CEL = new CQGCEL();
            m_CEL.DataError += new CQG._ICQGCELEvents_DataErrorEventHandler(CEL_DataError);
            m_CEL.DataConnectionStatusChanged += new CQG._ICQGCELEvents_DataConnectionStatusChangedEventHandler(CEL_DataConnectionStatusChanged);
            m_CEL.HistoricalSessionsResolved += new CQG._ICQGCELEvents_HistoricalSessionsResolvedEventHandler(CEL_HistoricalSessionsResolved);
            m_CEL.APIConfiguration.ReadyStatusCheck = eReadyStatusCheck.rscOff;
            m_CEL.APIConfiguration.CollectionsThrowException = false;
            m_CEL.APIConfiguration.TimeZoneCode = eTimeZone.tzCentral;
            // Disables the controls
            CEL_DataConnectionStatusChanged(eConnectionStatus.csConnectionDown);
            // Starts up the CQGCEL
            m_CEL.Startup();
        }

        private void CEL_DataError(object cqg_error, string error_description)
        {

            if (cqg_error is CQGError)
            {
                CQGError cqgErr = (CQGError)cqg_error;

                if (cqgErr.Code == 102)
                {
                    error_description += " Restart the application.";
                }
                else if (cqgErr.Code == 125)
                {
                    error_description += " Turn on CQG Client and restart the application.";
                }
            }
        }
        private void CEL_DataConnectionStatusChanged(CQG.eConnectionStatus new_status)
        {

            string sInfo;

            if (new_status == eConnectionStatus.csConnectionUp)
            {
                sInfo = "DATA Connection is UP";
            }
            else if (new_status == eConnectionStatus.csConnectionDelayed)
            {
                sInfo = "DATA Connection is Delayed";
            }
            else
            {
                sInfo = "DATA Connection is Down";
            }

        }

        private void CEL_HistoricalSessionsResolved(CQGSessionsCollection sessColl,
                                                  CQGHistoricalSessionsRequest request,
                                                  CQGError error)
        {
            string status = (error != null) ? "Failed" : "Succeeded";
            // Dump all data
            DumpAllData(sessColl);
        }

        private void DumpAllData(CQGSessionsCollection historicalSessions)
        {

            long sessIndex = 1;
            long holidayIndex = 1;

            // Dump all data
            foreach (CQGSessions sessions in historicalSessions)
            {
                DumpSessions(sessions, sessIndex);
                DumpHolidays(sessions.Holidays, holidayIndex);

                // Count current indices for sessions and holidays
                sessIndex += sessions.Count;
                holidayIndex += sessions.Holidays.Count;
            }
        }

        private void DumpSessions(CQGSessions sessions, long index)
        {
            foreach (CQGSession session in sessions)
            {
                histSession = new Dictionary<string, string>();
                histSession.Add("Index", (index).ToString());
                histSession.Add("Name", session.Name);
                histSession.Add("Number", session.Number.ToString());
                histSession.Add("Activation Date", GetValueAsString(session.ActivationDate, true));
                histSession.Add("Start Time", session.StartTime.ToShortTimeString());
                histSession.Add("End Time", session.EndTime.ToShortTimeString());
                histSession.Add("Is Primary", session.PrimaryFlag ? "Yes" : "No");
                histSession.Add("Type", session.Type.ToString());
                histSession.Add("Working Week Days", GetSessionWorkingDays(session.WorkingWeekDays));
                histSession.Add("Description Number", sessions.DescriptionNumber.ToString());
                histSession.Add("Description Start", GetValueAsString(sessions.DescriptionStart, true));
                histSession.Add("Description End", GetValueAsString(sessions.DescriptionEnd, true));

                histSessions.Add(histSession);

                ++index;
            }
        }

        private void DumpHolidays(CQGHolidays holidays, long index)
        {
            foreach (CQGHoliday holiday in holidays)
            {
                histSession = new Dictionary<string, string>();
                histSession.Add("Index", (index).ToString());
                histSession.Add("HolidayDate", GetValueAsString(holiday.HolidayDate, false));
                histSession.Add("SessionMask", "0x" + holiday.SessionMask.ToString());
                histSession.Add("IsDaily", holiday.IsDaily ? "TRUE" : "FALSE");
                histSession.Add("SessDescNumber", holidays.Sessions.DescriptionNumber.ToString());
                histSession.Add("SessDescStart", GetValueAsString(holidays.Sessions.DescriptionStart, true));
                histSession.Add("SessDescEnd", GetValueAsString(holidays.Sessions.DescriptionEnd, true));

                histSessions.Add(histSession);

                ++index;
            }
        }

        private static string GetSessionWorkingDays(eSessionWeekDays weekDay)
        {
            string sResult;

            sResult = (((weekDay & eSessionWeekDays.swdSunday) == eSessionWeekDays.swdSunday) ? "S" : "-").ToString();
            sResult += (((weekDay & eSessionWeekDays.swdMonday) == eSessionWeekDays.swdMonday) ? "M" : "-").ToString();
            sResult += (((weekDay & eSessionWeekDays.swdTuesday) == eSessionWeekDays.swdTuesday) ? "T" : "-").ToString();
            sResult += (((weekDay & eSessionWeekDays.swdWednesday) == eSessionWeekDays.swdWednesday) ? "W" : "-").ToString();
            sResult += (((weekDay & eSessionWeekDays.swdThursday) == eSessionWeekDays.swdThursday) ? "T" : "-").ToString();
            sResult += (((weekDay & eSessionWeekDays.swdFriday) == eSessionWeekDays.swdFriday) ? "F" : "-").ToString();
            sResult += (((weekDay & eSessionWeekDays.swdSaturday) == eSessionWeekDays.swdSaturday) ? "S" : "-").ToString();

            return sResult;
        }

        private string GetValueAsString(object val, bool withTime)
        {
            string sResult = N_A;

            try
            {
                if (m_CEL.IsValid(val))
                {
                    if (val.GetType().FullName == "System.DateTime")
                    {
                        if (withTime)
                        {
                            sResult = System.Convert.ToDateTime(val).ToString("g");
                        }
                        else
                        {
                            sResult = System.Convert.ToDateTime(val).ToString("d");
                        }
                    }
                    else
                    {
                        sResult = val.ToString();
                    }
                }
            }
            catch
            {
                sResult = N_A;
                return sResult;
            }

            return sResult;
        }


    }
}
