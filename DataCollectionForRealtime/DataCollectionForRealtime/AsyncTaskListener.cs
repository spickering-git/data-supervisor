using System;

namespace DataCollectionForRealtime
{
    /// <summary>
    /// This helper class carries three roles for asynchronous tasks:
    /// 1. The reporter of messages.
    /// 2. The reporter of progress.
    /// 3. The measurer and reporter of the "Records Per Second" quantity (RPS).
    ///    (The class contains logic of the RPS indicator located on DB form.)
    /// </summary>
    static class AsyncTaskListener
    {
        public delegate void UpdateDelegate(string msg = null);
        public static event UpdateDelegate Updated;

        public static void LogMessage(string msg)
        {
            // Update text box
            Updated.Invoke(msg);
        }

        public static void LogMessageFormat(string msgPat, params object[] args)
        {
            // Update text box
            Updated.Invoke(string.Format(msgPat, args));
        }
    }
}
