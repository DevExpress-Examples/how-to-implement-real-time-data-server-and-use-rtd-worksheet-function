using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TestExcelRTDServer {   
    [
        Guid("B6AF4673-200B-413c-8536-1F778AC14DE1"),
        ProgId("My.Sample.RtdServer"),
        ComVisible(true)
    ]
    public class RtdServer : IRtdServer {
        private IRTDUpdateEvent m_callback;
        private Timer m_timer;
        private Dictionary<int, string> m_topics;
        private static Random random = new Random();

        public int ServerStart(IRTDUpdateEvent callback) {
            m_callback = callback;

            m_timer = new Timer();
            m_timer.Tick += new EventHandler(TimerEventHandler);
            m_timer.Interval = 500;

            m_topics = new Dictionary<int, string>();

            return 1;
        }

        public void ServerTerminate() {
            if (null != m_timer) {
                m_timer.Dispose();
                m_timer = null;
            }
        }

        public object ConnectData(int topicId,
                                  ref Array strings,
                                  ref bool newValues) {
            if (1 != strings.Length) {
                return "Exactly one parameter is required";
            }

            string value = strings.GetValue(0).ToString();

            m_topics[topicId] = value;
            m_timer.Start();
            return GetNextValue(value);
        }

        public void DisconnectData(int topicId) {
            m_topics.Remove(topicId);
        }

        public Array RefreshData(ref int topicCount) {
            object[,] data = new object[2, m_topics.Count];

            int index = 0;

            foreach (int topicId in m_topics.Keys) {
                data[0, index] = topicId;
                data[1, index] = GetNextValue(m_topics[topicId]);

                ++index;
            }

            topicCount = m_topics.Count;

            m_timer.Start();
            return data;
        }

        public int Heartbeat() {
            return 1;
        }

        private void TimerEventHandler(object sender,
                                       EventArgs args) {
            m_timer.Stop();
            m_callback.UpdateNotify();
        }

        private static double GetNextValue(string value) {
            double quote;
            switch (value)
            {
                case "MSFT":
                    quote = 40;
                    break;
                case "FB":
                    quote = 60;
                    break;
                case "YHOO":
                    quote = 36;
                    break;
                default:
                    quote = Double.Parse(value);
                    break;
            }
            return quote + (random.NextDouble() * 10.0 - 5.0);
        }
    }
}

namespace Microsoft.Office.Interop.Excel {
    [Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComImport()]
    [TypeIdentifier]
    [ComVisible(true)]
    public interface IRTDUpdateEvent {
        void UpdateNotify();

//        int HeartbeatInterval { get; set; }

//        void Disconnect();
    }
    [Guid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComImport()]
    [TypeIdentifier]
    [ComVisible(true)]
    public interface IRtdServer {
        [DispId(10)]
        int ServerStart(IRTDUpdateEvent callback);
        [DispId(11)]
        object ConnectData(int topicId,
                           [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array strings,
                           ref bool newValues);

        [DispId(12)]
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        Array RefreshData(ref int topicCount);

        [DispId(13)]
        void DisconnectData(int topicId);

        [DispId(14)]
        int Heartbeat();

        [DispId(15)]
        void ServerTerminate();
    }
}
