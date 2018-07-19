// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Plugin.cs" company="KriaSoft LLC">
//   Copyright © 2013 Konstantin Tarkus, KriaSoft LLC. See LICENSE.txt
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace AmiBroker.Plugin
{
    using System;
    using System.Collections.Specialized;
    using System.Diagnostics;
    using System.Drawing;
    using System.Linq;
    //using System.Net.Http;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Windows.Controls;
    using System.Windows.Forms;
    using Controls;
    using Models;
    using System.Collections.Generic;
    using System.Collections;

    using System.Net;
    using System.IO;

    // Timer
    using System.Threading;

    // JSON
    using Newtonsoft.Json;

    using RGiesecke.DllExport;

    /// <summary>
    /// Standard implementation of a typical AmiBroker plug-ins.
    /// </summary>
    public class Plugin
    {
        public static string webSiteURL = "http://amicoins.ru/version.txt";
        public static string pluginVersion;
        public static bool isUpdateAvailable = false;

        /// <summary>
        /// Plugin status code
        /// </summary>
        static StatusCode Status = StatusCode.OK;

        // Дескриптор окна = null по умолчанию
        static IntPtr mainWnd = IntPtr.Zero;

        static int refrashTime = -1;
        /// <summary>
        /// Default encoding
        /// </summary>
        static Encoding encoding = Encoding.GetEncoding("windows-1251"); // TODO: Update it based on your preferences

        static DataSource DataSource;

        // Timer
        static AutoResetEvent autoEvent = new AutoResetEvent(false);

        static TimerCallback timerCallback = new TimerCallback(NotifyStreamingUpdate);

        // Таймер инициализирован на тик раз в минуту!!!
        static System.Threading.Timer timer = new System.Threading.Timer(timerCallback, autoEvent, 0, 1000);

        /// <summary>
        /// WPF user control which is used to display right-click context menu.
        /// </summary>
        static RightClickMenu RightClickMenu;

        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static void GetPluginInfo(ref PluginInfo pluginInfo)
        {
            //MessageBox.Show("GetPluginInfo...");

            pluginInfo.Name = "CryptoCurrencies Data Plug-in (Demo)";
            pluginInfo.Vendor = "Arakcheev V.A.";
            pluginInfo.Type = PluginType.Data;
            
            pluginInfo.Version = 0915; // v0.9.13
            pluginVersion = "0.9.15";

            pluginInfo.IDCode = new PluginID("DEMO");
            pluginInfo.Certificate = 0;
            pluginInfo.MinAmiVersion = 5600000; // v5.60
            pluginInfo.StructSize = Marshal.SizeOf((PluginInfo)pluginInfo);
        }

 
        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static void Init()
        {
           // MessageBox.Show("Init...");
            // add version control
            string webVersion = MarketData.getMarketData(webSiteURL);

            // Если ошибка
            if (webVersion != "")
            {
                if (String.Compare(pluginVersion, webVersion) != 0)
                    isUpdateAvailable = true;
            }
        }

        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static void Release()
        {
            //MessageBox.Show("Release");
        }

        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static unsafe void Notify(PluginNotification* notification)
        {
            //MessageBox.Show("Notify...");

            switch (notification->Reason)
            {
                case PluginNotificationReason.DatabaseLoaded:
 
                    mainWnd = notification->MainWnd;
                    RightClickMenu = new RightClickMenu(notification->MainWnd);
                    break;

                case PluginNotificationReason.DatabaseUnloaded:
                    break;

                case PluginNotificationReason.StatusRightClick:

                    RightClickMenu.ContextMenu.IsOpen = true;
                    break;

                case PluginNotificationReason.SettingsChange:
                    break;
            }
        }

        /// <summary>
        /// GetQuotesEx function is functional equivalent fo GetQuotes but
        /// handles new Quotation format with 64 bit date/time stamp and floating point volume/open int
        /// and new Aux fields
        /// it also takes pointer to context that is reserved for future use (can be null)
        /// Called by AmiBroker 5.27 and above 
        /// </summary>
        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static unsafe int GetQuotesEx(string ticker, Periodicity periodicity, int lastValid, int size, Quotation* quotes, GQEContext* context)
        {

            // Статус - в ожидании данных
            Status = StatusCode.Wait;

            // при запуске из базы
            if (refrashTime == -1)
                refrashTime = (int)periodicity;

            // ticker - наименование пары
            // periodicity - период опроса (day, mminute, ...)
            
           string baseURL;
            
            // Если не фьючерсы Okex
           if (ticker.IndexOf("Future Contracts") < 0)
           {

               baseURL = "https://min-api.cryptocompare.com/data/";


               // Вроде как ограничение за последние 30 дней = 30 записей по умолчанию (макс 2000)
               if (periodicity == Periodicity.EndOfDay)
                   baseURL += "histoday?";

               // Вроде как ограничение за последние 7 дней = 168 записей по умолчанию (макс 2000)
               if (periodicity == Periodicity.OneHour)
                   baseURL += "histohour?";

               // Вроде как ограничение за последние 24 часа = 1440 записей по умолчанию (макс 2000)
               if (periodicity == Periodicity.OneMinute)
                   baseURL += "histominute?aggregate=1&";

               // 5 минутная агрегация
               if (periodicity == Periodicity.FiveMinutes)
                   baseURL += "histominute?aggregate=5&";

               // 15 минутная агрегация
               if (periodicity == Periodicity.FifteenMinutes)
                   baseURL += "histominute?aggregate=15&";

           }
           else
           { 
               // https://www.okex.com/api/v1/future_kline.do?symbol=btc_usd&contract_type=this_week&type=5min

               baseURL = "https://www.okex.com/api/v1/future_kline.do?type=";

               // Вроде как ограничение за последние 30 дней = 30 записей по умолчанию (макс 2000)
               if (periodicity == Periodicity.EndOfDay)
                   baseURL += "1day";

               // Вроде как ограничение за последние 7 дней = 168 записей по умолчанию (макс 2000)
               if (periodicity == Periodicity.OneHour)
                   baseURL += "1hour";

               // Вроде как ограничение за последние 24 часа = 1440 записей по умолчанию (макс 2000)
               if (periodicity == Periodicity.OneMinute)
                   baseURL += "1min";

               // 5 минутная агрегация
               if (periodicity == Periodicity.FiveMinutes)
                   baseURL += "5min";

               // 15 минутная агрегация
               if (periodicity == Periodicity.FifteenMinutes)
                   baseURL += "15min";

           }
            // 2. Выбираем маркет и пару

            // Разбор вида [1]/[2] - [3]
            // [1] - 1й символ
            // [2] - 2й символ
            // [3] - маркет

            try
            {

                string s1 = ticker.Substring(0, ticker.IndexOf("/"));
                string s2 = ticker.Substring(ticker.IndexOf("/") + 1, ticker.IndexOf(" ") - ticker.IndexOf("/"));
                string market = ticker.Substring(ticker.IndexOf(" - ") + 3);

                // Коррекция CryptoCompare
                if (market.IndexOf("CryptoCompare") > -1)
                    market = "CCCAGG";


                if (ticker.IndexOf("Future Contracts") < 0)
                {
                    // Добавим пары и маркет в URL
                    baseURL += "fsym=" + s1.Trim() + "&tsym=" + s2.Trim() + "&e=" + market.Trim() + "&allData=true";
                }
                else
                { 
                    //symbol=btc_usd
                    baseURL += "&symbol=" + s1.Trim().ToLower() + "_" + s2.Trim().ToLower() + "&contract_type=" + RightClickMenu.contractType + "&size=" + size;
                }

            }
            catch (Exception e)
            {
                Status = StatusCode.Error;
                Log.Write("The symbol translate error!");

                return lastValid + 1;
            }

             string jsonResponse = MarketData.getMarketData(baseURL);

                  // Если ошибка
             if (String.IsNullOrEmpty(jsonResponse))
                 {
                     Status = StatusCode.Error;
                     Log.Write("Error: " + MarketData.getLastError());
                     return lastValid + 1;
                 }

                 MarketData.hasData = false;

                 // В jsonResponse имеем файл ответа сервера                
                // Парсим
                List<Quotation> newQuotes = null;                

                if (ticker.IndexOf("Future Contracts") < 0)
                {
                    HistoryData data = null;                    

                    try
                    {
                        data = JsonConvert.DeserializeObject<HistoryData>(jsonResponse);
                    }
                    catch (Exception e)
                    {
                        Log.Write("Json deserialize error: " + e.Message);
                        return lastValid + 1;
                    }

                    // В итоге имеем данные
                    newQuotes = data.GetQuotesList();
                }
                else
                {
                    IList<IList<string>> data = null;

                    try
                    {
                        data = JsonConvert.DeserializeObject<IList<IList<string>>>(jsonResponse);
                    }
                    catch (Exception e)
                    {
                        Log.Write("Json deserialize error: " + e.Message);
                        return lastValid + 1;
                    }
    
                    foreach (IList<string> item in data)
                    {
                        Quotation qt = new Quotation();

                        AmiDate tickerDate = new AmiDate(HistoryData.UnixTimeStampToDateTime(ulong.Parse(item[0]) / 1000));

                        try
                        {
                            qt.Open = float.Parse(item[1].Replace(".", ","));
                            qt.High = float.Parse(item[2].Replace(".", ","));
                            qt.Low = float.Parse(item[3].Replace(".", ","));
                            qt.Price = float.Parse(item[4].Replace(".", ","));
                            qt.Volume = float.Parse(item[5].Replace(".", ","));

                        }
                        catch (Exception e)
                        { 
                            Log.Write(e.Message);
                        }
                      
                        qt.DateTime = tickerDate.ToUInt64();
                        

                        newQuotes.Add(qt);
                    }                  

                }

               
                // Запись данных в БД

                #region Пишем с нуля - БД не заполнена
                if (lastValid < 0)
                {
                    lastValid = 0;

                    // Что меньше: длинна полученных данных или Limit?
                    int count = Math.Min(size, newQuotes.Count);
                    int i = 0, j = 0;

                    for (i = 0; i < count; i++)
                    {
                        j = newQuotes.Count - count + i;

                        // Date
                        quotes[i].DateTime = newQuotes[j].DateTime;

                        // ticker
                        quotes[i].Open = newQuotes[j].Open;
                        quotes[i].High = newQuotes[j].High;
                        quotes[i].Low = newQuotes[j].Low;
                        quotes[i].Price = newQuotes[j].Price;
                        quotes[i].Volume = newQuotes[j].Volume;

                        // Extra data
                        quotes[i].OpenInterest = 0;
                        quotes[i].AuxData1 = 0;
                        quotes[i].AuxData2 = 0;

                        lastValid++;
                    }

                    // Меняем статус на ОК
                    if (!isUpdateAvailable)
                        Status = StatusCode.OK;
                    else
                        Status = StatusCode.Update;

                    return lastValid;
                }

                #endregion

                #region Кастрация данных до последнего тайминга
                
                for (var i = 0; i < newQuotes.Count; i++)
                {
                    // Сейчас у нас в массиве только свежак
                    if (newQuotes[i].DateTime < quotes[lastValid].DateTime)
                    {
                        newQuotes.RemoveAt(i);
                        i--;
                    }
                }                

                #endregion

                // Если нечего показывать - выход
                if (newQuotes.Count == 0)
                    return lastValid + 1;

                #region Данные уже есть в БД - в корридор входим

                if (newQuotes.Count <= (size - lastValid - 1))
                {
                    foreach (Quotation item in newQuotes)
                    {

                        // Date
                        quotes[lastValid].DateTime = item.DateTime;

                        // ticker
                        quotes[lastValid].Open = item.Open;
                        quotes[lastValid].High = item.High;
                        quotes[lastValid].Low = item.Low;
                        quotes[lastValid].Price = item.Price;
                        quotes[lastValid].Volume = item.Volume;

                        // Extra data
                        quotes[lastValid].OpenInterest = 0;
                        quotes[lastValid].AuxData1 = 0;
                        quotes[lastValid].AuxData2 = 0;

                        lastValid++;                    
                    }

                    // Меняем статус на ОК
                    if (!isUpdateAvailable)
                        Status = StatusCode.OK;
                    else
                        Status = StatusCode.Update;

                    return lastValid;
                }

                #endregion

                #region Данные уже есть в БД - в корридор не входим - сдвиг массива

                    lastValid = 0;

                    // Смещение первой части
                    while (lastValid < (size - newQuotes.Count))
                    {

                        quotes[lastValid].DateTime = quotes[lastValid + newQuotes.Count].DateTime;
                        quotes[lastValid].Open = quotes[lastValid + newQuotes.Count].Open;
                        quotes[lastValid].High = quotes[lastValid + newQuotes.Count].High;
                        quotes[lastValid].Low = quotes[lastValid + newQuotes.Count].Low;
                        quotes[lastValid].Price = quotes[lastValid + newQuotes.Count].Price;
                        quotes[lastValid].Volume = quotes[lastValid + newQuotes.Count].Volume;

                        lastValid++;
                    }                    

                    // КОпируем остатки
                    foreach (Quotation item in newQuotes)
                    {

                        quotes[lastValid].DateTime = item.DateTime;
                        quotes[lastValid].Open = item.Open;
                        quotes[lastValid].High = item.High;
                        quotes[lastValid].Low = item.Low;
                        quotes[lastValid].Price = item.Price;
                        quotes[lastValid].Volume = item.Volume;

                        lastValid++;
                    }

                     // Меняем статус на ОК
                     if (!isUpdateAvailable)
                         Status = StatusCode.OK;
                     else
                         Status = StatusCode.Update;

                    return lastValid;

                #endregion    

        }

        


        public unsafe delegate void* Alloc(uint size);

        ///// <summary>
        ///// GetExtra data is optional function for retrieving non-quotation data
        ///// </summary>
        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static AmiVar GetExtraData(string ticker, string name, int arraySize, Periodicity periodicity, Alloc alloc)
        {
            return new AmiVar();
        }

        /// <summary>
        /// GetSymbolLimit function is optional, used only by real-time plugins
        /// </summary>
        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static int GetSymbolLimit()
        {
            return 10000;
        }

        /// <summary>
        /// GetStatus function is optional, used mostly by few real-time plugins
        /// </summary>
        /// <param name="statusPtr">A pointer to <see cref="AmiBrokerPlugin.PluginStatus"/></param>
        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static void GetStatus(IntPtr statusPtr)
        {
            switch (Status)
            {
                case StatusCode.OK:
                    SetStatus(statusPtr, StatusCode.OK, Color.LightGreen, "OK", "CryptoCurrencies Data Plug-in is running...");
                    break;
                case StatusCode.Wait:
                    SetStatus(statusPtr, StatusCode.Wait, Color.LightBlue, "WAIT", "Retrieving data...");
                    break;
                case StatusCode.Error:
                    SetStatus(statusPtr, StatusCode.Error, Color.Red, "ERR", "An error occured");
                    break;
                case StatusCode.Update:
                    SetStatus(statusPtr, StatusCode.Update, Color.LightSeaGreen, "Update", "Plugin update available at http://amicoins.ru");
                    break;
                default:
                    SetStatus(statusPtr, StatusCode.Unknown, Color.LightGray, "Ukno", "Unknown status");
                    break;
            }
        }

        #region Helper Functions

        // Устанавливает полное имя криптовалютной пары
        static void UpdateFullName(IntPtr ptr, string fullName)
        {
            var si = (StockInfo)Marshal.PtrToStructure(ptr, typeof(StockInfo));

            if (fullName != null)
            {
                var enc = Encoding.GetEncoding("windows-1251");
                var bytes = enc.GetBytes(fullName);

                for (var i = 0; i < (bytes.Length > 127 ? 127 : bytes.Length); i++)
                {
                    #if (_WIN64)
                        Marshal.WriteByte(new IntPtr(ptr.ToInt64() + 144 + i), bytes[i]);
                    #else
                        Marshal.WriteByte(new IntPtr(ptr.ToInt32() + 144 + i), bytes[i]);
                    #endif
                }

                #if (_WIN64)
                    Marshal.WriteByte(new IntPtr(ptr.ToInt64() + 144 + bytes.Length), 0x0);
                #else
                    Marshal.WriteByte(new IntPtr(ptr.ToInt32() + 144 + bytes.Length), 0x0);
                #endif
            }
        }


        // Добавляет в маркет сток
        private static void addStock(string stockName, int marketIndex = 0, string fullName = "")
        {
            IntPtr stock;

            stock = AddStockNew(stockName);

            // index of market

            //Marshal.WriteInt32(new IntPtr(stock.ToInt32() + 476), marketIndex);
            Marshal.WriteInt64(new IntPtr(stock.ToInt64() + 476), marketIndex);          

            // Update fullName
            UpdateFullName(stock, fullName);
        }

        /// <summary>
        /// Configure function is called when user presses "Configure" button in File->Database Settings
        /// </summary>
        /// <param name="path">Path to AmiBroker database</param>
        /// <param name="site">A pointer to <see cref="AmiBrokerPlugin.InfoSite"/></param>
        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static int Configure(string path, IntPtr infoSitePtr)
        {
            Status = StatusCode.Wait;

            // 32 bit
            // GetStockQty = (GetStockQtyDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt32(new IntPtr(infoSitePtr.ToInt32() + 4))), typeof(GetStockQtyDelegate));
            
            // 64 bit
            //GetStockQty = (GetStockQtyDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt32(new IntPtr(infoSitePtr.ToInt32() + 8))), typeof(GetStockQtyDelegate));

            #if (_WIN64)
                // 64 bit
                SetCategoryName = (SetCategoryNameDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt64(new IntPtr(infoSitePtr.ToInt64() + 24))), typeof(SetCategoryNameDelegate));
            #else
                // 32 bit
                SetCategoryName = (SetCategoryNameDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt32(new IntPtr(infoSitePtr.ToInt32() + 12))), typeof(SetCategoryNameDelegate));
            #endif

            //GetCategoryName = (GetCategoryNameDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt32(new IntPtr(infoSitePtr.ToInt32() + 30))), typeof(GetCategoryNameDelegate));
            //SetIndustrySector = (SetIndustrySectorDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt32(new IntPtr(infoSitePtr.ToInt32() + 16))), typeof(SetIndustrySectorDelegate));
            //GetIndustrySector = (GetIndustrySectorDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt32(new IntPtr(infoSitePtr.ToInt32() + 20))), typeof(GetIndustrySectorDelegate));

            #if (_WIN64)
                // 64 bit
                AddStockNew = (AddStockNewDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt64(new IntPtr(infoSitePtr.ToInt64() + 56))), typeof(AddStockNewDelegate));
            #else
                // 32 bit
                AddStockNew = (AddStockNewDelegate)Marshal.GetDelegateForFunctionPointer(new IntPtr(Marshal.ReadInt32(new IntPtr(infoSitePtr.ToInt32() + 28))), typeof(AddStockNewDelegate));
            #endif

            SetCategoryName(0, 0, "Bitfinex");
            SetCategoryName(0, 1, "Bitstamp");
            SetCategoryName(0, 2, "Binance");
            SetCategoryName(0, 3, "BitTrex");            
            SetCategoryName(0, 4, "Coinbase");
            SetCategoryName(0, 5, "CryptoCompare");
            SetCategoryName(0, 6, "Top 50 CryptoCompare");
            SetCategoryName(0, 7, "Okex");
            SetCategoryName(0, 8, "Okex Futures");
            SetCategoryName(0, 9, "Poloniex");
            SetCategoryName(0, 10, "Kraken");
            SetCategoryName(0, 11, "BX.in.th");
            SetCategoryName(0, 12, "WavesDEX");
            SetCategoryName(0, 13, "BTCMarkets");


            for (var i = 0; i < StockList.stockItems.Count; i++)
                addStock(StockList.stockItems[i].shotName, StockList.stockItems[i].marketIndex, StockList.stockItems[i].fullName);
            
            Status = StatusCode.OK;

            MessageBox.Show("Configure is completed!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

           // Log.Write("Configure complete...");

            return 1;
        }

        /// <summary>
        /// SetTimeBase function is called when user is changing base time interval in File->Database Settings
        /// </summary>
        [DllExport(CallingConvention = CallingConvention.Cdecl)]
        public static int SetTimeBase(int periodicity)
        {
           // MessageBox.Show("SetTimeBase...");

            switch (periodicity)
            { 
                case (int)Periodicity.OneMinute:
                case (int)Periodicity.FiveMinutes:
                case (int)Periodicity.FifteenMinutes:
                case (int)Periodicity.OneHour:
                case (int)Periodicity.EndOfDay:

                    refrashTime = periodicity;

                    return 1;            
            }

            return 0;
            
            //return periodicity >= (int)Periodicity.OneHour && periodicity <= (int)Periodicity.EndOfDay ? 1 : 0;
        }

        /// <summary>
        /// Notify AmiBroker that new streaming data arrived
        /// </summary>
        static void NotifyStreamingUpdate(Object stateInfo)
        {

            //MessageBox.Show("NotifyStreamingUpdate...");

            // Если нет инициализации - не обновляем!
            if (mainWnd == IntPtr.Zero)
                return;

            DateTime curTime = DateTime.UtcNow;

            // Минуты - обновление каждую минуту хх:01 сек
            if ((refrashTime == (int)Periodicity.OneMinute) && (curTime.Second == 1))
            {
                NativeMethods.SendMessage(mainWnd, 0x0400 + 13000, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            // 5 Минут - обновление каждую минуту хх:01 сек при % 5 == 0
            if ((refrashTime == (int)Periodicity.FiveMinutes) && (curTime.Second == 1) && (curTime.Minute % 5 == 0))
            {
                NativeMethods.SendMessage(mainWnd, 0x0400 + 13000, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            // 15 Минут - обновление каждую минуту хх:01 сек при % 5 == 0
            if ((refrashTime == (int)Periodicity.FifteenMinutes) && (curTime.Second == 1) && (curTime.Minute % 15 == 0))
            {
                NativeMethods.SendMessage(mainWnd, 0x0400 + 13000, IntPtr.Zero, IntPtr.Zero);
                return;
            }

             // 1 час - обновление если минуты = 0 и 01 сек
            if ((refrashTime == (int)Periodicity.OneHour) && (curTime.Second == 1) && (curTime.Minute == 0))
            {
                NativeMethods.SendMessage(mainWnd, 0x0400 + 13000, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            // Ежедневное обновление             
            if ((refrashTime == (int)Periodicity.EndOfDay) && (curTime.Second == 5) && (curTime.Minute == 5) && (curTime.Hour == 0))
            {
                NativeMethods.SendMessage(mainWnd, 0x0400 + 13000, IntPtr.Zero, IntPtr.Zero);
                return;
            }            
                        
        }

        /// <summary>
        /// Update status of the plugin
        /// </summary>
        /// <param name="statusPtr">A pointer to <see cref="AmiBrokerPlugin.PluginStatus"/></param>
        static void SetStatus(IntPtr statusPtr, StatusCode code, Color color, string shortMessage, string fullMessage)
        {

            #if (_WIN64) 
                Marshal.WriteInt64(new IntPtr(statusPtr.ToInt64() + 4), (int)code);
                Marshal.WriteInt64(new IntPtr(statusPtr.ToInt64() + 8), color.R);
                Marshal.WriteInt64(new IntPtr(statusPtr.ToInt64() + 9), color.G);
                Marshal.WriteInt64(new IntPtr(statusPtr.ToInt64() + 10), color.B);
            #else
                Marshal.WriteInt32(new IntPtr(statusPtr.ToInt32() + 4), (int)code);
                Marshal.WriteInt32(new IntPtr(statusPtr.ToInt32() + 8), color.R);
                Marshal.WriteInt32(new IntPtr(statusPtr.ToInt32() + 9), color.G);
                Marshal.WriteInt32(new IntPtr(statusPtr.ToInt32() + 10), color.B);
            #endif


            var msg = encoding.GetBytes(fullMessage);

            for (int i = 0; i < (msg.Length > 255 ? 255 : msg.Length); i++)
            {
                #if (_WIN64)
                    Marshal.WriteInt64(new IntPtr(statusPtr.ToInt64() + 12 + i), msg[i]);
                #else
                    Marshal.WriteInt32(new IntPtr(statusPtr.ToInt32() + 12 + i), msg[i]);
                #endif 
            }

            #if (_WIN64)
                Marshal.WriteInt64(new IntPtr(statusPtr.ToInt64() + 12 + msg.Length), 0x0);
            #else
                 Marshal.WriteInt32(new IntPtr(statusPtr.ToInt32() + 12 + msg.Length), 0x0);
            #endif

            msg = encoding.GetBytes(shortMessage);

            for (int i = 0; i < (msg.Length > 31 ? 31 : msg.Length); i++)
            {
                #if (_WIN64)
                    Marshal.WriteInt64(new IntPtr(statusPtr.ToInt64() + 268 + i), msg[i]);
                #else
                    Marshal.WriteInt32(new IntPtr(statusPtr.ToInt32() + 268 + i), msg[i]);
                #endif

            }
      
            #if (_WIN64)
                Marshal.WriteInt64(new IntPtr(statusPtr.ToInt64() + 268 + msg.Length), 0x0);
            #else
                Marshal.WriteInt32(new IntPtr(statusPtr.ToInt32() + 268 + msg.Length), 0x0);
            #endif

        }

        #endregion

        #region AmiBroker Method Delegates
                        
        [UnmanagedFunctionPointer(CallingConvention.Cdecl, SetLastError = true)]
        delegate int GetStockQtyDelegate();

        private static GetStockQtyDelegate GetStockQty;

        [UnmanagedFunctionPointer(CallingConvention.Cdecl, SetLastError = true)]
        delegate int SetCategoryNameDelegate(int category, int item, string name);

        private static SetCategoryNameDelegate SetCategoryName;

        [UnmanagedFunctionPointer(CallingConvention.Cdecl, SetLastError = true)]
        delegate string GetCategoryNameDelegate(int category, int item);

        private static GetCategoryNameDelegate GetCategoryName;

        [UnmanagedFunctionPointer(CallingConvention.Cdecl, SetLastError = true)]
        delegate int SetIndustrySectorDelegate(int industry, int sector);

        private static SetIndustrySectorDelegate SetIndustrySector;

        [UnmanagedFunctionPointer(CallingConvention.Cdecl, SetLastError = true)]
        delegate int GetIndustrySectorDelegate(int industry);

        private static GetIndustrySectorDelegate GetIndustrySector;

        // Only available if called from AmiBroker 5.27 or higher
        [UnmanagedFunctionPointer(CallingConvention.Cdecl, SetLastError = true)]
        public delegate IntPtr AddStockNewDelegate([MarshalAs(UnmanagedType.LPStr)] string ticker);

        private static AddStockNewDelegate AddStockNew;


        #endregion
    } 
}
