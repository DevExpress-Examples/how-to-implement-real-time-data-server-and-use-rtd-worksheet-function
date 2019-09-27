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
        IRTDUpdateEvent m_callback;
        Timer m_timer;
        Dictionary<int, TopicData> m_topics;
        Dictionary<string, CompanyData> companies;

        public RtdServer() {
            this.companies = new Dictionary<string, CompanyData>();
            this.companies.Add("MSFT", new CompanyData("MSFT", 40, 176));
            this.companies.Add("FB", new CompanyData("FB", 60, 210));
            this.companies.Add("YHOO", new CompanyData("YHOO", 36, 54));
            this.companies.Add("NOK", new CompanyData("NOK", 50, 100));
            UpdatePrices();
        }

        public int ServerStart(IRTDUpdateEvent callback) {
            m_callback = callback;

            m_timer = new Timer();
            m_timer.Tick += new EventHandler(TimerEventHandler);
            m_timer.Interval = 500;

            m_topics = new Dictionary<int, TopicData>();

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
            if (2 != strings.Length)
                return "Two parameters are required";
            string symbol = strings.GetValue(0).ToString();
            string type = strings.GetValue(1).ToString();

            TopicData data = new TopicData(symbol, type);

            m_topics[topicId] = data;
            m_timer.Start();
            return GetNextValue(data);
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

        void TimerEventHandler(object sender,
                                       EventArgs args) {
            m_timer.Stop();
            UpdatePrices();
            m_callback.UpdateNotify();
        }

        object GetNextValue(TopicData data) {
            CompanyData companyData;
            if (companies.TryGetValue(data.Symbol, out companyData)) {
                switch (data.Type) {
                    case "PRICE":
                        return companyData.Price;
                    case "CHANGE":
                        return companyData.Change;
                    case "SHARES":
                        return companyData.Shares;
                }
            }
            return "#Invalid";
        }

        void UpdatePrices() {
            foreach (CompanyData companyData in companies.Values)
                companyData.UpdatePrice();
        }
    }
}

public class TopicData {
    public TopicData(string symbol, string type) {
        Symbol = symbol;
        Type = type;
    }

    public string Symbol { get; }
    public string Type { get; }
}

public class CompanyData {
    static Random random = new Random();

    public CompanyData(string symbol, double quote, int shares) {
        this.Symbol = symbol;
        this.Quote = quote;
        this.Shares = shares;
    }

    double Quote { get; }
    public string Symbol { get; }
    public int Shares { get; }
    public double Price { get; private set; }
    public double Change { get; private set; }

    public void UpdatePrice() {
        this.Price = Quote + (random.NextDouble() * 10.0 - 5.0);
        this.Change = (this.Price - this.Quote) / this.Quote;
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
