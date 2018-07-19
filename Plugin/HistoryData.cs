using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using AmiBroker.Plugin.Models;

/*
    {
        "Response":"Success",
        "Message": "Battlecruiser operational.Make it happen.Set a course.Take it slow.",
        "Type":100,
        "Aggregated":false,
        "Data":[{
		            "time": 1445040000.0,
		            "open": 263.01,
		            "high": 274.94,
		            "low": 262.5,
		            "close": 270.51,
		            "volumefrom": 129992.56434794325,
		            "volumeto": 35024476.597084962
	            },
                ....
               ],
        "FirstValueInArray": true,
        "TimeTo": 1453075200,
        "TimeFrom": 1445040000
    }
  
*/
namespace AmiBroker.Plugin
{
    #region Types
    // {"time":1417132800,"close":376.28,"high":381.34,"low":360.57,"open":363.59,"volumefrom":8617.15,"volumeto":3220878.18}
    public class Ticker
    {
        public ulong time { get; set; }
        public float close { get; set; }
        public float high { get; set; }
        public float low { get; set; }
        public float open { get; set; }
        public float volumefrom { get; set; }
        public float volumeto { get; set; }
    }
    #endregion

    #region Types
    class HistoryData
    {
        public string Response { get; set; }
        public string Message { get; set; }
        public short Type { get; set; }
        public bool Aggregated { get; set; }
        public IList<Ticker> Data { get; set; }
        public bool FirstValueInArray { get; set; }
        public ulong TimeTo { get; set; }
        public ulong TimeFrom { get; set; }

        public static DateTime UnixTimeStampToDateTime(ulong unixTimeStamp)
        {
            // Unix timestamp is seconds past epoch
            System.DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            dtDateTime = dtDateTime.AddSeconds(unixTimeStamp).ToUniversalTime(); // ToLocalTime();
            return dtDateTime;
        }

        public List<Quotation> GetQuotesList()
        {
            List<Quotation> result = new List<Quotation>();

            foreach (Ticker item in this.Data)
            {
                Quotation qt = new Quotation();

                qt.High = item.high;
                qt.Low = item.low;
                qt.Open = item.open;
                qt.Price = item.close;
                qt.Volume = item.volumefrom; ///item.volumeto - item.volumefrom;

                // Date
                AmiDate tickerDate = new AmiDate(HistoryData.UnixTimeStampToDateTime(item.time));
                qt.DateTime = tickerDate.ToUInt64();

                result.Add(qt);
            }

            return result;
        }

        public Quotation[] GetQuotes()
        {
            // TODO: Return the list of quotes for the specified ticker.

            List<Quotation> QList = new List<Quotation>();

            foreach (Ticker item in this.Data)
            {
                Quotation qt = new Quotation();

                qt.High = item.high;
                qt.Low = item.low;
                qt.Open = item.open;
                qt.Price = item.close;
                qt.Volume = item.volumefrom; ///item.volumeto - item.volumefrom;

                // Date
                AmiDate tickerDate = new AmiDate(HistoryData.UnixTimeStampToDateTime(item.time));              
                qt.DateTime = tickerDate.ToUInt64();

                QList.Add(qt);
            }

            return QList.ToArray();
        }

    }
    #endregion
}
