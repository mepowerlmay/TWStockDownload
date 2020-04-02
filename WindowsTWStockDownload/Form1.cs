using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Dapper;
using System.Threading;
using WindowsTWStockDownload.Model;
using System.Net.Security;

namespace WindowsTWStockDownload
{
    public partial class Form1 : Form
    {

        private Thread thread;

        /// <summary>
        /// 是否轉檔
        /// </summary>
        bool IsTrans;

        public Form1()
        {
           

            InitializeComponent();

        }


        private void Form1_Load(object sender, EventArgs e)
        {
        
        }



        #region 跨執行續處理


        private void UpdatedtpSDate()
        {
        //    DateTime date = DateTime.Now;
            //var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            //dtpSDate.Value = firstDayOfMonth;
        }

        delegate void delUpdateTxtRrsult(string txt);


        public void UpdateTxtRrsult(string txt) {

            txtResult.Text = txt;
            txtResult.Update();
        }


        #endregion

//        Public Function ValidateServerCertificate(sender As Object, certification As System.Security.Cryptography.X509Certificates.X509Certificate, _
//                                              chain As System.Security.Cryptography.X509Certificates.X509Chain, sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
//        Return True
//End Function



       //     public bool 

        public void TransStock() {

            DateTime date = DateTime.Now;
            //var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);


            MethodInvoker mi = new MethodInvoker(this.UpdatedtpSDate);
            this.BeginInvoke(mi, null);
            //   int TimeRange = (dtpEDate.Value - dtpSDate.Value).Days;
            Dictionary<string, string> arryStocks = GetStocks(); // 股票清單

            IsTrans = true;

            //   Thread.Sleep(1);//每(1/1000秒)取得一次狀態,使用時須先導入using System.Threading;
            SpinWait.SpinUntil(() => false, 1);
            Application.DoEvents();//將狀態停止並顯示出來

            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            //   Thread t = new Thread();

            foreach (var item in arryStocks.Keys)
            {
                //Thread.Sleep(1);//每(1/1000秒)取得一次狀態,使用時須先導入using System.Threading;
                SpinWait.SpinUntil(() => false, 1);
                Application.DoEvents();//將狀態停止並顯示出來

                //WebClient

                try
                {
                    string targetUrl = string.Format(@"https://www.tse.com.tw/exchangeReport/STOCK_DAY?response=json&date={0}&stockNo={1}", dtpSDate.Value.ToString("yyyyMM01"), item);

                    Dictionary<string, object> dics = GetStockData(targetUrl);

                    if (dics["stat"].ToString().StartsWith("很抱歉"))
                    {
                        continue;
                    }

                    if (dics["stat"].ToString().StartsWith("查詢日期大於今日"))
                    {
                        continue;
                    }

                    if (dics["stat"].ToString().StartsWith("查詢日期小於"))
                    {
                        continue;
                    }

                    var obj = JsonConvert.DeserializeObject(dics["data"].ToString());
                    JArray JStockData = (JArray)obj;

                    // mi = new MethodInvoker(this.UpdateTxtRrsult);




                  string result = string.Format("股號:{0} 轉檔月份:{1} 資料比數:{2}", item, dtpSDate.Value.ToString("yyyyMM"), JStockData.Count)
                        + Environment.NewLine;



                    delUpdateTxtRrsult d = new delUpdateTxtRrsult(UpdateTxtRrsult);
                    this.BeginInvoke(d, result);

                


                    SaveStock(JStockData, item);
                }
                catch (Exception ex)
                {
                    //休息2小時  1秒*60*2
                    //  System.Threading.Thread.Sleep(1000 * 60 * 120); ;


                    SpinWait.SpinUntil(() => false, 1000 * 60 * 120);
                    // throw ex;
                }

                if (!IsTrans)
                {
                    break;

                }


                GC.Collect();
                SpinWait.SpinUntil(() => false, 3500);
                //System.Threading.Thread.Sleep(3500);
            }

            thread.Abort();
            
            thread = null;

        }

        private void btnTransfer_Click(object sender, EventArgs e)
        {

             thread = new Thread(TransStock);
            thread.IsBackground = true;
            thread.Start();







        }

        /// <summary>
        /// 股票存檔
        /// </summary>
        /// <param name="jStockData"></param>
        /// <param name="stock"></param>
        private void SaveStock(JArray jStockData,string stock)
        {
            string strconn = @" data source= 127.0.0.1; Initial Catalog = DBStock; User Id = sa; Password = 321456852; ";
            IDbConnection conn = new SqlConnection(strconn);
         

            foreach (JToken item in jStockData)
            {
                //Thread.Sleep(1);//每(1/1000秒)取得一次狀態,使用時須先導入using System.Threading;

                //SpinWait.SpinUntil(() => false, 1);
                Application.DoEvents();//將狀態停止並顯示出來

                string tempdate = item[0].ToString().Replace("/", "");

                string sqlselect = string.Format("select * from Twstock where chDate = '{0}' and StockCode = '{1}'", tempdate, stock);

                IEnumerable<Twstock> GroupTwstockPreData = conn.Query<Twstock>(sqlselect);              

                if (GroupTwstockPreData.Count() > 0)
                {
                    GroupTwstockPreData = null;
                    continue;
                }
                

                Twstock model = new Twstock();


        
                model.chDate = tempdate;


                string tempClosePrice = item[6].ToString();
                if (tempClosePrice == "--")
                {
                    model.ClosePrice = 0;
                }
                else
                {
                    model.ClosePrice = Convert.ToDouble(tempClosePrice);
                }


           

                model.UpOrDown = item[7].ToString(); 
                model.TradeCount= item[1].ToString();
                model.TradeMoney= item[2].ToString();

                model.OpenPrice= item[3].ToString();
                if (model.OpenPrice == "--")
                {
                    model.OpenPrice = "0";
                }
                model.HightPrice= item[4].ToString();
                if (model.HightPrice == "--")
                {
                    model.HightPrice = "0";
                }
                model.LowPrice = item[5].ToString();
                if (model.LowPrice == "--")
                {
                    model.LowPrice = "0";
                }


                model.StockCode = stock;

                string insertsql = @"insert into Twstock (chDate
                                                         ,ClosePrice
                                                         ,UpOrDown
                                                         ,TradeCount
                                                         ,TradeMoney
                                                         ,OpenPrice
                                                         ,HightPrice
                                                         ,LowPrice
                                                         ,StockCode) VALUES (
                                                          @chDate
                                                         ,@ClosePrice
                                                         ,@UpOrDown
                                                         ,@TradeCount
                                                         ,@TradeMoney
                                                         ,@OpenPrice
                                                         ,@HightPrice
                                                         ,@LowPrice
                                                         ,@StockCode )";

                conn.Execute(insertsql, model);
             string    result = string.Concat("正在寫入 股號" + stock+ "   " + model.chDate) + Environment.NewLine;



                delUpdateTxtRrsult d = new delUpdateTxtRrsult(UpdateTxtRrsult);
                this.BeginInvoke(d, result);



                Application.DoEvents();
            }
        }

        //Public Function ValidateServerCertificate(sender As Object, certification As System.Security.Cryptography.X509Certificates.X509Certificate, _
        //chain As System.Security.Cryptography.X509Certificates.X509Chain, sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        //        Return True
        //End Function


        public bool ValidateServerCertificate(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certification, System.Security.Cryptography.X509Certificates.X509Chain Chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }



        /// <summary>
        /// 取得股票資料
        /// </summary>
        /// <param name="targetUrl"></param>
        /// <returns></returns>
        private Dictionary<string, object> GetStockData(string targetUrl)
        {

            HttpWebRequest request = HttpWebRequest.Create(targetUrl) as HttpWebRequest;
            //request.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
            //request.se
            request.Method = "GET";
            request.ContentType = "application/x-www-form-urlencoded";


            string jsondata = "";
            // 取得回應資料
            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                {
                    jsondata = sr.ReadToEnd();
                }
            }

            SpinWait.SpinUntil(() => false, 5*1000);

          //  System.Threading.Thread.Sleep(5000);

            var settings = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            Dictionary<string, object> values = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsondata, settings);

            return values;
        }


        /// <summary>
        /// 取得目前股票清單
        /// </summary>
        /// <returns></returns>
        private Dictionary<string,string> GetStocks() {

         var dics = new Dictionary<string, string>();

            dics.Add("1101", "台泥");
            dics.Add("1102", "亞泥");
            dics.Add("1103", "嘉泥");
            dics.Add("1104", "環泥");
            dics.Add("1108", "幸福");
            dics.Add("1109", "信大");
            dics.Add("1110", "東泥");
            dics.Add("1201", "味全");
            dics.Add("1203", "味王");
            dics.Add("1210", "大成");
            dics.Add("1213", "大飲");
            dics.Add("1215", "卜蜂");
            dics.Add("1216", "統一");
            dics.Add("1217", "愛之味");
            dics.Add("1218", "泰山");
            dics.Add("1219", "福壽");
            dics.Add("1220", "台榮");
            dics.Add("1225", "福懋油");
            dics.Add("1227", "佳格");
            dics.Add("1229", "聯華");
            dics.Add("1231", "聯華食");
            dics.Add("1232", "大統益");
            dics.Add("1233", "天仁");
            dics.Add("1234", "黑松");
            dics.Add("1235", "興泰");
            dics.Add("1236", "宏亞");
            dics.Add("1256", "鮮活果汁-KY");
            dics.Add("1262", "綠悅-KY");
            dics.Add("1301", "台塑");
            dics.Add("1303", "南亞");
            dics.Add("1304", "台聚");
            dics.Add("1305", "華夏");
            dics.Add("1307", "三芳");
            dics.Add("1308", "亞聚");
            dics.Add("1309", "台達化");
            dics.Add("1310", "台苯");
            dics.Add("1312", "國喬");
            dics.Add("1313", "聯成");
            dics.Add("1314", "中石化");
            dics.Add("1315", "達新");
            dics.Add("1316", "上曜");
            dics.Add("1319", "東陽");
            dics.Add("1321", "大洋");
            dics.Add("1323", "永裕");
            dics.Add("1324", "地球");
            dics.Add("1325", "恆大");
            dics.Add("1326", "台化");
            dics.Add("1337", "再生-KY");
            dics.Add("1338", "廣華-KY");
            dics.Add("1339", "昭輝");
            dics.Add("1340", "勝悅-KY");
            dics.Add("1341", "富林-KY");
            dics.Add("1402", "遠東新");
            dics.Add("1409", "新纖");
            dics.Add("1410", "南染");
            dics.Add("1413", "宏洲");
            dics.Add("1414", "東和");
            dics.Add("1416", "廣豐");
            dics.Add("1417", "嘉裕");
            dics.Add("1418", "東華");
            dics.Add("1419", "新紡");
            dics.Add("1423", "利華");
            dics.Add("1432", "大魯閣");
            dics.Add("1434", "福懋");
            dics.Add("1435", "中福");
            dics.Add("1436", "華友聯");
            dics.Add("1437", "勤益控");
            dics.Add("1438", "裕豐");
            dics.Add("1439", "中和");
            dics.Add("1440", "南紡");
            dics.Add("1441", "大東");
            dics.Add("1442", "名軒");
            dics.Add("1443", "立益");
            dics.Add("1444", "力麗");
            dics.Add("1445", "大宇");
            dics.Add("1446", "宏和");
            dics.Add("1447", "力鵬");
            dics.Add("1449", "佳和");
            dics.Add("1451", "年興");
            dics.Add("1452", "宏益");
            dics.Add("1453", "大將");
            dics.Add("1454", "台富");
            dics.Add("1455", "集盛");
            dics.Add("1456", "怡華");
            dics.Add("1457", "宜進");
            dics.Add("1459", "聯發");
            dics.Add("1460", "宏遠");
            dics.Add("1463", "強盛");
            dics.Add("1464", "得力");
            dics.Add("1465", "偉全");
            dics.Add("1466", "聚隆");
            dics.Add("1467", "南緯");
            dics.Add("1468", "昶和");
            dics.Add("1470", "大統新創");
            dics.Add("1471", "首利");
            dics.Add("1472", "三洋紡");
            dics.Add("1473", "台南");
            dics.Add("1474", "弘裕");
            dics.Add("1475", "本盟");
            dics.Add("1476", "儒鴻");
            dics.Add("1477", "聚陽");
            dics.Add("1503", "士電");
            dics.Add("1504", "東元");
            dics.Add("1506", "正道");
            dics.Add("1507", "永大");
            dics.Add("1512", "瑞利");
            dics.Add("1513", "中興電");
            dics.Add("1514", "亞力");
            dics.Add("1515", "力山");
            dics.Add("1516", "川飛");
            dics.Add("1517", "利奇");
            dics.Add("1519", "華城");
            dics.Add("1521", "大億");
            dics.Add("1522", "堤維西");
            dics.Add("1524", "耿鼎");
            dics.Add("1525", "江申");
            dics.Add("1526", "日馳");
            dics.Add("1527", "鑽全");
            dics.Add("1528", "恩德");
            dics.Add("1529", "樂士");
            dics.Add("1530", "亞崴");
            dics.Add("1531", "高林股");
            dics.Add("1532", "勤美");
            dics.Add("1533", "車王電");
            dics.Add("1535", "中宇");
            dics.Add("1536", "和大");
            dics.Add("1537", "廣隆");
            dics.Add("1538", "正峰新");
            dics.Add("1539", "巨庭");
            dics.Add("1540", "喬福");
            dics.Add("1541", "錩泰");
            dics.Add("1558", "伸興");
            dics.Add("1560", "中砂");
            dics.Add("1568", "倉佑");
            dics.Add("1582", "信錦");
            dics.Add("1583", "程泰");
            dics.Add("1587", "吉茂");
            dics.Add("1589", "永冠-KY");
            dics.Add("1590", "亞德客-KY");
            dics.Add("1592", "英瑞-KY");
            dics.Add("1598", "岱宇");
            dics.Add("1603", "華電");
            dics.Add("1604", "聲寶");
            dics.Add("1605", "華新");
            dics.Add("1608", "華榮");
            dics.Add("1609", "大亞");
            dics.Add("1611", "中電");
            dics.Add("1612", "宏泰");
            dics.Add("1614", "三洋電");
            dics.Add("1615", "大山");
            dics.Add("1616", "億泰");
            dics.Add("1617", "榮星");
            dics.Add("1618", "合機");
            dics.Add("1626", "艾美特-KY");
            dics.Add("1701", "中化");
            dics.Add("1702", "南僑");
            dics.Add("1707", "葡萄王");
            dics.Add("1708", "東鹼");
            dics.Add("1709", "和益");
            dics.Add("1710", "東聯");
            dics.Add("1711", "永光");
            dics.Add("1712", "興農");
            dics.Add("1713", "國化");
            dics.Add("1714", "和桐");
            dics.Add("1717", "長興");
            dics.Add("1718", "中纖");
            dics.Add("1720", "生達");
            dics.Add("1721", "三晃");
            dics.Add("1722", "台肥");
            dics.Add("1723", "中碳");
            dics.Add("1724", "台硝");
            dics.Add("1725", "元禎");
            dics.Add("1726", "永記");
            dics.Add("1727", "中華化");
            dics.Add("1730", "花仙子");
            dics.Add("1731", "美吾華");
            dics.Add("1732", "毛寶");
            dics.Add("1733", "五鼎");
            dics.Add("1734", "杏輝");
            dics.Add("1735", "日勝化");
            dics.Add("1736", "喬山");
            dics.Add("1737", "臺鹽");
            dics.Add("1760", "寶齡富錦");
            dics.Add("1762", "中化生");
            dics.Add("1773", "勝一");
            dics.Add("1776", "展宇");
            dics.Add("1783", "和康生");
            dics.Add("1786", "科妍");
            dics.Add("1789", "神隆");
            dics.Add("1802", "台玻");
            dics.Add("1805", "寶徠");
            dics.Add("1806", "冠軍");
            dics.Add("1808", "潤隆");
            dics.Add("1809", "中釉");
            dics.Add("1810", "和成");
            dics.Add("1817", "凱撒衛");
            dics.Add("1902", "台紙");
            dics.Add("1903", "士紙");
            dics.Add("1904", "正隆");
            dics.Add("1905", "華紙");
            dics.Add("1906", "寶隆");
            dics.Add("1907", "永豐餘");
            dics.Add("1909", "榮成");
            dics.Add("2002", "中鋼");
            dics.Add("2006", "東和鋼鐵");
            dics.Add("2007", "燁興");
            dics.Add("2008", "高興昌");
            dics.Add("2009", "第一銅");
            dics.Add("2010", "春源");
            dics.Add("2012", "春雨");
            dics.Add("2013", "中鋼構");
            dics.Add("2014", "中鴻");
            dics.Add("2015", "豐興");
            dics.Add("2017", "官田鋼");
            dics.Add("2020", "美亞");
            dics.Add("2022", "聚亨");
            dics.Add("2023", "燁輝");
            dics.Add("2024", "志聯");
            dics.Add("2025", "千興");
            dics.Add("2027", "大成鋼");
            dics.Add("2028", "威致");
            dics.Add("2029", "盛餘");
            dics.Add("2030", "彰源");
            dics.Add("2031", "新光鋼");
            dics.Add("2032", "新鋼");
            dics.Add("2033", "佳大");
            dics.Add("2034", "允強");
            dics.Add("2038", "海光");
            dics.Add("2049", "上銀");
            dics.Add("2059", "川湖");
            dics.Add("2062", "橋椿");
            dics.Add("2069", "運錩");
            dics.Add("2101", "南港");
            dics.Add("2102", "泰豐");
            dics.Add("2103", "台橡");
            dics.Add("2104", "國際中橡");
            dics.Add("2105", "正新");
            dics.Add("2106", "建大");
            dics.Add("2107", "厚生");
            dics.Add("2108", "南帝");
            dics.Add("2109", "華豐");
            dics.Add("2114", "鑫永銓");
            dics.Add("2115", "六暉-KY");
            dics.Add("2201", "裕隆");
            dics.Add("2204", "中華");
            dics.Add("2206", "三陽工業");
            dics.Add("2207", "和泰車");
            dics.Add("2208", "台船");
            dics.Add("2227", "裕日車");
            dics.Add("2228", "劍麟");
            dics.Add("2231", "為升");
            dics.Add("2236", "百達-KY");
            dics.Add("2239", "英利-KY");
            dics.Add("2243", "宏旭-KY");
            dics.Add("2301", "光寶科");
            dics.Add("2302", "麗正");
            dics.Add("2303", "聯電");
            dics.Add("2305", "全友");
            dics.Add("2308", "台達電");
            dics.Add("2312", "金寶");
            dics.Add("2313", "華通");
            dics.Add("2314", "台揚");
            dics.Add("2316", "楠梓電");
            dics.Add("2317", "鴻海");
            dics.Add("2321", "東訊");
            dics.Add("2323", "中環");
            dics.Add("2324", "仁寶");
            dics.Add("2327", "國巨");
            dics.Add("2328", "廣宇");
            dics.Add("2329", "華泰");
            dics.Add("2330", "台積電");
            dics.Add("2331", "精英");
            dics.Add("2332", "友訊");
            dics.Add("2337", "旺宏");
            dics.Add("2338", "光罩");
            dics.Add("2340", "光磊");
            dics.Add("2342", "茂矽");
            dics.Add("2344", "華邦電");
            dics.Add("2345", "智邦");
            dics.Add("2347", "聯強");
            dics.Add("2348", "海悅");
            dics.Add("2349", "錸德");
            dics.Add("2351", "順德");
            dics.Add("2352", "佳世達");
            dics.Add("2353", "宏碁");
            dics.Add("2354", "鴻準");
            dics.Add("2355", "敬鵬");
            dics.Add("2356", "英業達");
            dics.Add("2357", "華碩");
            dics.Add("2358", "廷鑫");
            dics.Add("2359", "所羅門");
            dics.Add("2360", "致茂");
            dics.Add("2362", "藍天");
            dics.Add("2363", "矽統");
            dics.Add("2364", "倫飛");
            dics.Add("2365", "昆盈");
            dics.Add("2367", "燿華");
            dics.Add("2368", "金像電");
            dics.Add("2369", "菱生");
            dics.Add("2371", "大同");
            dics.Add("2373", "震旦行");
            dics.Add("2374", "佳能");
            dics.Add("2375", "智寶");
            dics.Add("2376", "技嘉");
            dics.Add("2377", "微星");
            dics.Add("2379", "瑞昱");
            dics.Add("2380", "虹光");
            dics.Add("2382", "廣達");
            dics.Add("2383", "台光電");
            dics.Add("2385", "群光");
            dics.Add("2387", "精元");
            dics.Add("2388", "威盛");
            dics.Add("2390", "云辰");
            dics.Add("2392", "正崴");
            dics.Add("2393", "億光");
            dics.Add("2395", "研華");
            dics.Add("2397", "友通");
            dics.Add("2399", "映泰");
            dics.Add("2401", "凌陽");
            dics.Add("2402", "毅嘉");
            dics.Add("2404", "漢唐");
            dics.Add("2405", "浩鑫");
            dics.Add("2406", "國碩");
            dics.Add("2408", "南亞科");
            dics.Add("2409", "友達");
            dics.Add("2412", "中華電");
            dics.Add("2413", "環科");
            dics.Add("2414", "精技");
            dics.Add("2415", "錩新");
            dics.Add("2417", "圓剛");
            dics.Add("2419", "仲琦");
            dics.Add("2420", "新巨");
            dics.Add("2421", "建準");
            dics.Add("2423", "固緯");
            dics.Add("2424", "隴華");
            dics.Add("2425", "承啟");
            dics.Add("2426", "鼎元");
            dics.Add("2427", "三商電");
            dics.Add("2428", "興勤");
            dics.Add("2429", "銘旺科");
            dics.Add("2430", "燦坤");
            dics.Add("2431", "聯昌");
            dics.Add("2433", "互盛電");
            dics.Add("2434", "統懋");
            dics.Add("2436", "偉詮電");
            dics.Add("2438", "翔耀");
            dics.Add("2439", "美律");
            dics.Add("2440", "太空梭");
            dics.Add("2441", "超豐");
            dics.Add("2442", "新美齊");
            dics.Add("2443", "新利虹");
            dics.Add("2444", "兆勁");
            dics.Add("2448", "晶電");
            dics.Add("2449", "京元電子");
            dics.Add("2450", "神腦");
            dics.Add("2451", "創見");
            dics.Add("2453", "凌群");
            dics.Add("2454", "聯發科");
            dics.Add("2455", "全新");
            dics.Add("2456", "奇力新");
            dics.Add("2457", "飛宏");
            dics.Add("2458", "義隆");
            dics.Add("2459", "敦吉");
            dics.Add("2460", "建通");
            dics.Add("2461", "光群雷");
            dics.Add("2462", "良得電");
            dics.Add("2464", "盟立");
            dics.Add("2465", "麗臺");
            dics.Add("2466", "冠西電");
            dics.Add("2467", "志聖");
            dics.Add("2468", "華經");
            dics.Add("2471", "資通");
            dics.Add("2472", "立隆電");
            dics.Add("2474", "可成");
            dics.Add("2475", "華映");
            dics.Add("2476", "鉅祥");
            dics.Add("2477", "美隆電");
            dics.Add("2478", "大毅");
            dics.Add("2480", "敦陽科");
            dics.Add("2481", "強茂");
            dics.Add("2482", "連宇");
            dics.Add("2483", "百容");
            dics.Add("2484", "希華");
            dics.Add("2485", "兆赫");
            dics.Add("2486", "一詮");
            dics.Add("2488", "漢平");
            dics.Add("2489", "瑞軒");
            dics.Add("2491", "吉祥全");
            dics.Add("2492", "華新科");
            dics.Add("2493", "揚博");
            dics.Add("2495", "普安");
            dics.Add("2496", "卓越");
            dics.Add("2497", "怡利電");
            dics.Add("2498", "宏達電");
            dics.Add("2499", "東貝");
            dics.Add("2501", "國建");
            dics.Add("2504", "國產");
            dics.Add("2505", "國揚");
            dics.Add("2506", "太設");
            dics.Add("2509", "全坤建");
            dics.Add("2511", "太子");
            dics.Add("2514", "龍邦");
            dics.Add("2515", "中工");
            dics.Add("2516", "新建");
            dics.Add("2520", "冠德");
            dics.Add("2524", "京城");
            dics.Add("2527", "宏璟");
            dics.Add("2528", "皇普");
            dics.Add("2530", "華建");
            dics.Add("2534", "宏盛");
            dics.Add("2535", "達欣工");
            dics.Add("2536", "宏普");
            dics.Add("2537", "聯上發");
            dics.Add("2538", "基泰");
            dics.Add("2539", "櫻花建");
            dics.Add("2540", "愛山林");
            dics.Add("2542", "興富發");
            dics.Add("2543", "皇昌");
            dics.Add("2545", "皇翔");
            dics.Add("2546", "根基");
            dics.Add("2547", "日勝生");
            dics.Add("2548", "華固");
            dics.Add("2597", "潤弘");
            dics.Add("2601", "益航");
            dics.Add("2603", "長榮");
            dics.Add("2605", "新興");
            dics.Add("2606", "裕民");
            dics.Add("2607", "榮運");
            dics.Add("2608", "嘉里大榮");
            dics.Add("2609", "陽明");
            dics.Add("2610", "華航");
            dics.Add("2611", "志信");
            dics.Add("2612", "中航");
            dics.Add("2613", "中櫃");
            dics.Add("2614", "東森");
            dics.Add("2615", "萬海");
            dics.Add("2616", "山隆");
            dics.Add("2617", "台航");
            dics.Add("2618", "長榮航");
            dics.Add("2630", "亞航");
            dics.Add("2633", "台灣高鐵");
            dics.Add("2634", "漢翔");
            dics.Add("2636", "台驊投控");
            dics.Add("2637", "慧洋-KY");
            dics.Add("2642", "宅配通");
            dics.Add("2701", "萬企");
            dics.Add("2702", "華園");
            dics.Add("2704", "國賓");
            dics.Add("2705", "六福");
            dics.Add("2706", "第一店");
            dics.Add("2707", "晶華");
            dics.Add("2712", "遠雄來");
            dics.Add("2722", "夏都");
            dics.Add("2723", "美食-KY");
            dics.Add("2727", "王品");
            dics.Add("2731", "雄獅");
            dics.Add("2739", "寒舍");
            dics.Add("2748", "雲品");
            dics.Add("2801", "彰銀");
            dics.Add("2809", "京城銀");
            dics.Add("2812", "台中銀");
            dics.Add("2816", "旺旺保");
            dics.Add("2820", "華票");
            dics.Add("2823", "中壽");
            dics.Add("2832", "台產");
            dics.Add("2834", "臺企銀");
            dics.Add("2836", "高雄銀");
            dics.Add("2838", "聯邦銀");
            dics.Add("2841", "台開");
            dics.Add("2845", "遠東銀");
            dics.Add("2849", "安泰銀");
            dics.Add("2850", "新產");
            dics.Add("2851", "中再保");
            dics.Add("2852", "第一保");
            dics.Add("2855", "統一證");
            dics.Add("2867", "三商壽");
            dics.Add("2880", "華南金");
            dics.Add("2881", "富邦金");
            dics.Add("2882", "國泰金");
            dics.Add("2883", "開發金");
            dics.Add("2884", "玉山金");
            dics.Add("2885", "元大金");
            dics.Add("2886", "兆豐金");
            dics.Add("2887", "台新金");
            dics.Add("2888", "新光金");
            dics.Add("2889", "國票金");
            dics.Add("2890", "永豐金");
            dics.Add("2891", "中信金");
            dics.Add("2892", "第一金");
            dics.Add("2897", "王道銀行");
            dics.Add("2901", "欣欣");
            dics.Add("2903", "遠百");
            dics.Add("2904", "匯僑");
            dics.Add("2905", "三商");
            dics.Add("2906", "高林");
            dics.Add("2908", "特力");
            dics.Add("2910", "統領");
            dics.Add("2911", "麗嬰房");
            dics.Add("2912", "統一超");
            dics.Add("2913", "農林");
            dics.Add("2915", "潤泰全");
            dics.Add("2923", "鼎固-KY");
            dics.Add("2929", "淘帝-KY");
            dics.Add("2936", "客思達-KY");
            dics.Add("2939", "凱羿-KY");
            dics.Add("3002", "歐格");
            dics.Add("3003", "健和興");
            dics.Add("3004", "豐達科");
            dics.Add("3005", "神基");
            dics.Add("3006", "晶豪科");
            dics.Add("3008", "大立光");
            dics.Add("3010", "華立");
            dics.Add("3011", "今皓");
            dics.Add("3013", "晟銘電");
            dics.Add("3014", "聯陽");
            dics.Add("3015", "全漢");
            dics.Add("3016", "嘉晶");
            dics.Add("3017", "奇鋐");
            dics.Add("3018", "同開");
            dics.Add("3019", "亞光");
            dics.Add("3021", "鴻名");
            dics.Add("3022", "威強電");
            dics.Add("3023", "信邦");
            dics.Add("3024", "憶聲");
            dics.Add("3025", "星通");
            dics.Add("3026", "禾伸堂");
            dics.Add("3027", "盛達");
            dics.Add("3028", "增你強");
            dics.Add("3029", "零壹");
            dics.Add("3030", "德律");
            dics.Add("3031", "佰鴻");
            dics.Add("3032", "偉訓");
            dics.Add("3033", "威健");
            dics.Add("3034", "聯詠");
            dics.Add("3035", "智原");
            dics.Add("3036", "文曄");
            dics.Add("3037", "欣興");
            dics.Add("3038", "全台");
            dics.Add("3040", "遠見");
            dics.Add("3041", "揚智");
            dics.Add("3042", "晶技");
            dics.Add("3043", "科風");
            dics.Add("3044", "健鼎");
            dics.Add("3045", "台灣大");
            dics.Add("3046", "建碁");
            dics.Add("3047", "訊舟");
            dics.Add("3048", "益登");
            dics.Add("3049", "和鑫");
            dics.Add("3050", "鈺德");
            dics.Add("3051", "力特");
            dics.Add("3052", "夆典");
            dics.Add("3054", "立萬利");
            dics.Add("3055", "蔚華科");
            dics.Add("3056", "總太");
            dics.Add("3057", "喬鼎");
            dics.Add("3058", "立德");
            dics.Add("3059", "華晶科");
            dics.Add("3060", "銘異");
            dics.Add("3062", "建漢");
            dics.Add("3090", "日電貿");
            dics.Add("3094", "聯傑");
            dics.Add("3130", "一零四");
            dics.Add("3149", "正達");
            dics.Add("3164", "景岳");
            dics.Add("3167", "大量");
            dics.Add("3189", "景碩");
            dics.Add("3209", "全科");
            dics.Add("3229", "晟鈦");
            dics.Add("3231", "緯創");
            dics.Add("3257", "虹冠電");
            dics.Add("3266", "昇陽");
            dics.Add("3296", "勝德");
            dics.Add("3305", "昇貿");
            dics.Add("3308", "聯德");
            dics.Add("3311", "閎暉");
            dics.Add("3312", "弘憶股");
            dics.Add("3321", "同泰");
            dics.Add("3338", "泰碩");
            dics.Add("3346", "麗清");
            dics.Add("3356", "奇偶");
            dics.Add("3376", "新日興");
            dics.Add("3380", "明泰");
            dics.Add("3383", "新世紀");
            dics.Add("3406", "玉晶光");
            dics.Add("3413", "京鼎");
            dics.Add("3416", "融程電");
            dics.Add("3419", "譁裕");
            dics.Add("3432", "台端");
            dics.Add("3437", "榮創");
            dics.Add("3443", "創意");
            dics.Add("3450", "聯鈞");
            dics.Add("3454", "晶睿");
            dics.Add("3481", "群創");
            dics.Add("3494", "誠研");
            dics.Add("3501", "維熹");
            dics.Add("3504", "揚明光");
            dics.Add("3515", "華擎");
            dics.Add("3518", "柏騰");
            dics.Add("3519", "綠能");
            dics.Add("3528", "安馳");
            dics.Add("3530", "晶相光");
            dics.Add("3532", "台勝科");
            dics.Add("3533", "嘉澤");
            dics.Add("3535", "晶彩科");
            dics.Add("3536", "誠創");
            dics.Add("3545", "敦泰");
            dics.Add("3550", "聯穎");
            dics.Add("3557", "嘉威");
            dics.Add("3576", "聯合再生");
            dics.Add("3579", "尚志");
            dics.Add("3583", "辛耘");
            dics.Add("3588", "通嘉");
            dics.Add("3591", "艾笛森");
            dics.Add("3593", "力銘");
            dics.Add("3596", "智易");
            dics.Add("3605", "宏致");
            dics.Add("3607", "谷崧");
            dics.Add("3617", "碩天");
            dics.Add("3622", "洋華");
            dics.Add("3645", "達邁");
            dics.Add("3653", "健策");
            dics.Add("3661", "世芯-KY");
            dics.Add("3665", "貿聯-KY");
            dics.Add("3669", "圓展");
            dics.Add("3673", "TPK-KY");
            dics.Add("3679", "新至陞");
            dics.Add("3682", "亞太電");
            dics.Add("3686", "達能");
            dics.Add("3694", "海華");
            dics.Add("3698", "隆達");
            dics.Add("3701", "大眾控");
            dics.Add("3702", "大聯大");
            dics.Add("3703", "欣陸");
            dics.Add("3704", "合勤控");
            dics.Add("3705", "永信");
            dics.Add("3706", "神達");
            dics.Add("3708", "上緯投控");
            dics.Add("3711", "日月光投控");
            dics.Add("3712", "永崴投控");
            dics.Add("4104", "佳醫");
            dics.Add("4106", "雃博");
            dics.Add("4108", "懷特");
            dics.Add("4119", "旭富");
            dics.Add("4133", "亞諾法");
            dics.Add("4137", "麗豐-KY");
            dics.Add("4141", "龍燈-KY");
            dics.Add("4142", "國光生");
            dics.Add("4144", "康聯-KY");
            dics.Add("4148 ", "全宇生技-KY");
            dics.Add("4155", "訊映");
            dics.Add("4164", "承業醫");
            dics.Add("4190", "佐登-KY");
            dics.Add("4306", "炎洲");
            dics.Add("4414", "如興");
            dics.Add("4426", "利勤");
            dics.Add("4438", "廣越");
            dics.Add("4526", "東台");
            dics.Add("4532", "瑞智");
            dics.Add("4536", "拓凱");
            dics.Add("4540", "全球傳動");
            dics.Add("4545", "銘鈺");
            dics.Add("4551", "智伸科");
            dics.Add("4552", "力達-KY");
            dics.Add("4555", "氣立");
            dics.Add("4557", "永新-KY");
            dics.Add("4560", "強信-KY");
            dics.Add("4562", "穎漢");
            dics.Add("4566", "時碩工業");
            dics.Add("4720", "德淵");
            dics.Add("4722", "國精化");
            dics.Add("4725", "信昌化");
            dics.Add("4737", "華廣");
            dics.Add("4739", "康普");
            dics.Add("4746", "台耀");
            dics.Add("4755", "三福化");
            dics.Add("4763", "材料-KY");
            dics.Add("4764", "雙鍵");
            dics.Add("4766", "南寶");
            dics.Add("4807", "日成-KY");
            dics.Add("4904", "遠傳");
            dics.Add("4906", "正文");
            dics.Add("4912", "聯德控股-KY");
            dics.Add("4915", "致伸");
            dics.Add("4916", "事欣科");
            dics.Add("4919", "新唐");
            dics.Add("4927", "泰鼎-KY");
            dics.Add("4930", "燦星網");
            dics.Add("4934", "太極");
            dics.Add("4935", "茂林-KY");
            dics.Add("4938", "和碩");
            dics.Add("4942", "嘉彰");
            dics.Add("4943", "康控-KY");
            dics.Add("4952", "凌通");
            dics.Add("4956", "光鋐");
            dics.Add("4958", "臻鼎-KY");
            dics.Add("4960", "誠美材");
            dics.Add("4961", "天鈺");
            dics.Add("4967", "十銓");
            dics.Add("4968", "立積");
            dics.Add("4976", "佳凌");
            dics.Add("4977", "眾達-KY");
            dics.Add("4989", "榮科");
            dics.Add("4994", "傳奇");
            dics.Add("4999", "鑫禾");
            dics.Add("5007", "三星");
            dics.Add("5203", "訊連");
            dics.Add("5215", "科嘉-KY");
            dics.Add("5225", "東科-KY");
            dics.Add("5234", "達興材料");
            dics.Add("5243", "乙盛-KY");
            dics.Add("5258", "虹堡");
            dics.Add("5259", "清惠");
            dics.Add("5264", "鎧勝-KY");
            dics.Add("5269", "祥碩");
            dics.Add("5284", "jpp-KY");
            dics.Add("5285", "界霖");
            dics.Add("5288", "豐祥-KY");
            dics.Add("5305", "敦南");
            dics.Add("5388", "中磊");
            dics.Add("5434", "崇越");
            dics.Add("5469", "瀚宇博");
            dics.Add("5471", "松翰");
            dics.Add("5484", "慧友");
            dics.Add("5515", "建國");
            dics.Add("5519", "隆大");
            dics.Add("5521", "工信");
            dics.Add("5522", "遠雄");
            dics.Add("5525", "順天");
            dics.Add("5531", "鄉林");
            dics.Add("5533", "皇鼎");
            dics.Add("5534", "長虹");
            dics.Add("5538", "東明-KY");
            dics.Add("5607", "遠雄港");
            dics.Add("5608", "四維航");
            dics.Add("5706", "鳳凰");
            dics.Add("5871", "中租-KY");
            dics.Add("5876", "上海商銀");
            dics.Add("5880", "合庫金");
            dics.Add("5906", "台南-KY");
            dics.Add("5907", "大洋-KY");
            dics.Add("6005", "群益證");
            dics.Add("6024", "群益期");
            dics.Add("6108", "競國");
            dics.Add("6112", "聚碩");
            dics.Add("6115", "鎰勝");
            dics.Add("6116", "彩晶");
            dics.Add("6117", "迎廣");
            dics.Add("6120", "達運");
            dics.Add("6128", "上福");
            dics.Add("6131", "悠克");
            dics.Add("6133", "金橋");
            dics.Add("6136", "富爾特");
            dics.Add("6139", "亞翔");
            dics.Add("6141", "柏承");
            dics.Add("6142", "友勁");
            dics.Add("6152", "百一");
            dics.Add("6153", "嘉聯益");
            dics.Add("6155", "鈞寶");
            dics.Add("6164", "華興");
            dics.Add("6165", "捷泰");
            dics.Add("6166", "凌華");
            dics.Add("6168", "宏齊");
            dics.Add("6172", "互億");
            dics.Add("6176", "瑞儀");
            dics.Add("6177", "達麗");
            dics.Add("6183", "關貿");
            dics.Add("6184", "大豐電");
            dics.Add("6189", "豐藝");
            dics.Add("6191", "精成科");
            dics.Add("6192", "巨路");
            dics.Add("6196", "帆宣");
            dics.Add("6197", "佳必琪");
            dics.Add("6201", "亞弘電");
            dics.Add("6202", "盛群");
            dics.Add("6205", "詮欣");
            dics.Add("6206", "飛捷");
            dics.Add("6209", "今國光");
            dics.Add("6213", "聯茂");
            dics.Add("6214", "精誠");
            dics.Add("6215", "和椿");
            dics.Add("6216", "居易");
            dics.Add("6224", "聚鼎");
            dics.Add("6225", "天瀚");
            dics.Add("6226", "光鼎");
            dics.Add("6230", "超眾");
            dics.Add("6235", "華孚");
            dics.Add("6239", "力成");
            dics.Add("6243", "迅杰");
            dics.Add("6251", "定穎");
            dics.Add("6257", "矽格");
            dics.Add("6269", "台郡");
            dics.Add("6271", "同欣電");
            dics.Add("6277", "宏正");
            dics.Add("6278", "台表科");
            dics.Add("6281", "全國電");
            dics.Add("6282", "康舒");
            dics.Add("6283", "淳安");
            dics.Add("6285", "啟碁");
            dics.Add("6288", "聯嘉");
            dics.Add("6289", "華上");
            dics.Add("6405", "悅城");
            dics.Add("6409", "旭隼");
            dics.Add("6412", "群電");
            dics.Add("6414", "樺漢");
            dics.Add("6415", "矽力-KY");
            dics.Add("6416", "瑞祺電通");
            dics.Add("6431", "光麗-KY");
            dics.Add("6442", "光聖");
            dics.Add("6443", "元晶");
            dics.Add("6449", "鈺邦");
            dics.Add("6451", "訊芯-KY");
            dics.Add("6452", "康友-KY");
            dics.Add("6456", "GIS-KY");
            dics.Add("6464", "台數科");
            dics.Add("6477", "安集");
            dics.Add("6504", "南六");
            dics.Add("6505", "台塑化");
            dics.Add("6525", "捷敏-KY");
            dics.Add("6531", "愛普");
            dics.Add("6533", "晶心科");
            dics.Add("6541", "泰福-KY");
            dics.Add("6552", "易華電");
            dics.Add("6558", "興能高");
            dics.Add("6573", "虹揚-KY");
            dics.Add("6579", "研揚");
            dics.Add("6581", "鋼聯");
            dics.Add("6582", "申豐");
            dics.Add("6591", "動力-KY");
            dics.Add("6605", "帝寶");
            dics.Add("6625", "必應");
            dics.Add("6641", "基士德-KY");
            dics.Add("6655", "科定");
            dics.Add("6666", "羅麗芬-KY");
            dics.Add("6668", "中揚光");
            dics.Add("6670", "復盛應用");
            dics.Add("6671", "三能-KY");
            dics.Add("6674", "鋐寶科技");
            dics.Add("8011", "台通");
            dics.Add("8016", "矽創");
            dics.Add("8021", "尖點");
            dics.Add("8028", "昇陽半導體");
            dics.Add("8033", "雷虎");
            dics.Add("8039", "台虹");
            dics.Add("8046", "南電");
            dics.Add("8070", "長華");
            dics.Add("8072", "陞泰");
            dics.Add("8081", "致新");
            dics.Add("8101", "華冠");
            dics.Add("8103", "瀚荃");
            dics.Add("8104", "錸寶");
            dics.Add("8105", "凌巨");
            dics.Add("8110", "華東");
            dics.Add("8112", "至上");
            dics.Add("8114", "振樺電");
            dics.Add("8131", "福懋科");
            dics.Add("8150", "南茂");
            dics.Add("8163", "達方");
            dics.Add("8201", "無敵");
            dics.Add("8210", "勤誠");
            dics.Add("8213", "志超");
            dics.Add("8215", "明基材");
            dics.Add("8222", "寶一");
            dics.Add("8249", "菱光");
            dics.Add("8261", "富鼎");
            dics.Add("8271", "宇瞻");
            dics.Add("8341", "日友");
            dics.Add("8367", "建新國際");
            dics.Add("8374", "羅昇");
            dics.Add("8404", "百和興業-KY");
            dics.Add("8411", "福貞-KY");
            dics.Add("8422", "可寧衛");
            dics.Add("8427", "基勝-KY");
            dics.Add("8429", "金麗-KY");
            dics.Add("8442", "威宏-KY");
            dics.Add("8443", "阿瘦");
            dics.Add("8454", "富邦媒");
            dics.Add("8463", "潤泰材");
            dics.Add("8464", "億豐");
            dics.Add("8466", "美喆-KY");
            dics.Add("8467", "波力-KY");
            dics.Add("8473", "山林水");
            dics.Add("8478", "東哥遊艇");
            dics.Add("8480", "泰昇-KY");
            dics.Add("8481", "政伸");
            dics.Add("8482", "商億-KY");
            dics.Add("8488", "吉源-KY");
            dics.Add("8497", "聯廣");
            dics.Add("8499", "鼎炫-KY");
            dics.Add("8926", "台汽電");
            dics.Add("8940", "新天地");
            dics.Add("8996", "高力");
            dics.Add("9802", "鈺齊-KY");
            dics.Add("9902", "台火");
            dics.Add("9904", "寶成");
            dics.Add("9905", "大華");
            dics.Add("9906", "欣巴巴");
            dics.Add("9907", "統一實");
            dics.Add("9908", "大台北");
            dics.Add("9910", "豐泰");
            dics.Add("9911", "櫻花");
            dics.Add("9912", "偉聯");
            dics.Add("9914", "美利達");
            dics.Add("9917", "中保");
            dics.Add("9918", "欣天然");
            dics.Add("9919", "康那香");
            dics.Add("9921", "巨大");
            dics.Add("9924", "福興");
            dics.Add("9925", "新保");
            dics.Add("9926", "新海");
            dics.Add("9927", "泰銘");
            dics.Add("9928", "中視");
            dics.Add("9929", "秋雨");
            dics.Add("9930", "中聯資源");
            dics.Add("9931", "欣高");
            dics.Add("9933", "中鼎");
            dics.Add("9934", "成霖");
            dics.Add("9935", "慶豐富");
            dics.Add("9937", "全國");
            dics.Add("9938", "百和");
            dics.Add("9939", "宏全");
            dics.Add("9940", "信義");
            dics.Add("9941", "裕融");
            dics.Add("9942", "茂順");
            dics.Add("9943", "好樂迪");
            dics.Add("9944", "新麗");
            dics.Add("9945", "潤泰新");
            dics.Add("9946", "三發地產");
            dics.Add("9955", "佳龍");
            dics.Add("9958", "世紀鋼");



            //上櫃

            dics.Add("1240", "茂生農經");
            dics.Add("1258", "其祥-KY");
            dics.Add("1259", "安心");
            dics.Add("1264", "德麥");
            dics.Add("1268", "漢來美食");
            dics.Add("1333", "恩得利");
            dics.Add("1336", "台翰");
            dics.Add("1565", "精華");
            dics.Add("1566", "捷邦");
            dics.Add("1569", "濱川");
            dics.Add("1570", "力肯");
            dics.Add("1580", "新麥");
            dics.Add("1584", "精剛");
            dics.Add("1586", "和勤");
            dics.Add("1591", "駿吉-KY");
            dics.Add("1593", "祺驊");
            dics.Add("1595", "川寶");
            dics.Add("1597", "直得");
            dics.Add("1599", "宏佳騰");
            dics.Add("1742", "台蠟");
            dics.Add("1752", "南光");
            dics.Add("1777", "生泰");
            dics.Add("1781", "合世");
            dics.Add("1784", "訊聯");
            dics.Add("1785", "光洋科");
            dics.Add("1787", "福盈科");
            dics.Add("1788", "杏昌");
            dics.Add("1795", "美時");
            dics.Add("1796", "金穎生技");
            dics.Add("1799", "易威");
            dics.Add("1813", "寶利徠");
            dics.Add("1815", "富喬");
            dics.Add("2035", "唐榮");
            dics.Add("2061", "風青");
            dics.Add("2063", "世鎧");
            dics.Add("2064", "晉椿");
            dics.Add("2065", "世豐");
            dics.Add("2066", "世德");
            dics.Add("2067", "嘉鋼");
            dics.Add("2070", "精湛");
            dics.Add("2221", "大甲");
            dics.Add("2230", "泰茂");
            dics.Add("2233", "宇隆");
            dics.Add("2235", "謚源");
            dics.Add("2596", "綠意");
            dics.Add("2640", "大車隊");
            dics.Add("2641", "正德");
            dics.Add("2643", "捷迅");
            dics.Add("2718", "晶悅");
            dics.Add("2719", "燦星旅");
            dics.Add("2724", "富驛-KY");
            dics.Add("2726", "雅茗-KY");
            dics.Add("2729", "瓦城");
            dics.Add("2732", "六角");
            dics.Add("2734", "易飛網");
            dics.Add("2736", "高野");
            dics.Add("2740", "天蔥");
            dics.Add("2745", "五福");
            dics.Add("2916", "滿心");
            dics.Add("2924", "東凌-KY");
            dics.Add("2926", "誠品生活");
            dics.Add("2928", "紅馬-KY");
            dics.Add("2937", "集雅社");
            dics.Add("3064", "泰偉");
            dics.Add("3066", "李洲");
            dics.Add("3067", "全域");
            dics.Add("3071", "協禧");
            dics.Add("3073", "凱柏實業");
            dics.Add("3078", "僑威");
            dics.Add("3081", "聯亞");
            dics.Add("3083", "網龍");
            dics.Add("3085", "新零售");
            dics.Add("3086", "華義");
            dics.Add("3088", "艾訊");
            dics.Add("3089", "元炬");
            dics.Add("3092", "鴻碩");
            dics.Add("3093", "港建");
            dics.Add("3095", "及成");
            dics.Add("3105", "穩懋");
            dics.Add("3114", "好德");
            dics.Add("3115", "寶島極");
            dics.Add("3118", "進階");
            dics.Add("3122", "笙泉");
            dics.Add("3128", "昇銳");
            dics.Add("3131", "弘塑");
            dics.Add("3141", "晶宏");
            dics.Add("3144", "新揚科");
            dics.Add("3147", "大綜");
            dics.Add("3152", "璟德");
            dics.Add("3162", "精確");
            dics.Add("3163", "波若威");
            dics.Add("3169", "亞信");
            dics.Add("3171", "新洲");
            dics.Add("3176", "基亞");
            dics.Add("3178", "公準");
            dics.Add("3188", "鑫龍騰");
            dics.Add("3191", "和進");
            dics.Add("3202", "樺晟");
            dics.Add("3205", "佰研");
            dics.Add("3206", "志豐");
            dics.Add("3207", "耀勝");
            dics.Add("3211", "順達");
            dics.Add("3213", "茂訊");
            dics.Add("3217", "優群");
            dics.Add("3218", "大學光");
            dics.Add("3219", "倚強");
            dics.Add("3221", "台嘉碩");
            dics.Add("3224", "三顧");
            dics.Add("3226", "至寶電");
            dics.Add("3227", "原相");
            dics.Add("3228", "金麗科");
            dics.Add("3230", "錦明");
            dics.Add("3232", "昱捷");
            dics.Add("3234", "光環");
            dics.Add("3236", "千如");
            dics.Add("3252", "海灣");
            dics.Add("3259", "鑫創");
            dics.Add("3260", "威剛");
            dics.Add("3264", "欣銓");
            dics.Add("3265", "台星科");
            dics.Add("3268", "海德威");
            dics.Add("3272", "東碩");
            dics.Add("3276", "宇環");
            dics.Add("3284", "太普高");
            dics.Add("3285", "微端");
            dics.Add("3287", "廣寰科");
            dics.Add("3288", "點晶");
            dics.Add("3289", "宜特");
            dics.Add("3290", "東浦");
            dics.Add("3293", "鈊象");
            dics.Add("3294", "英濟");
            dics.Add("3297", "杭特");
            dics.Add("3303", "岱稜");
            dics.Add("3306", "鼎天");
            dics.Add("3310", "佳穎");
            dics.Add("3313", "斐成");
            dics.Add("3317", "尼克森");
            dics.Add("3322", "建舜電");
            dics.Add("3323", "加百裕");
            dics.Add("3324", "雙鴻");
            dics.Add("3325", "旭品");
            dics.Add("3332", "幸康");
            dics.Add("3339", "泰谷");
            dics.Add("3354", "律勝");
            dics.Add("3360", "尚立");
            dics.Add("3362", "先進光");
            dics.Add("3363", "上詮");
            dics.Add("3372", "典範");
            dics.Add("3373", "熱映");
            dics.Add("3374", "精材");
            dics.Add("3379", "彬台");
            dics.Add("3388", "崇越電");
            dics.Add("3390", "旭軟");
            dics.Add("3402", "漢科");
            dics.Add("3426", "台興");
            dics.Add("3431", "長天");
            dics.Add("3434", "哲固");
            dics.Add("3438", "類比科");
            dics.Add("3441", "聯一光");
            dics.Add("3444", "利機");
            dics.Add("3452", "益通");
            dics.Add("3455", "由田");
            dics.Add("3465", "進泰電子");
            dics.Add("3466", "致振");
            dics.Add("3479", "安勤");
            dics.Add("3483", "力致");
            dics.Add("3484", "崧騰");
            dics.Add("3489", "森寶");
            dics.Add("3490", "單井");
            dics.Add("3491", "昇達科");
            dics.Add("3492", "長盛");
            dics.Add("3498", "陽程");
            dics.Add("3499", "環天科");
            dics.Add("3508", "位速");
            dics.Add("3511", "矽瑪");
            dics.Add("3512", "皇龍");
            dics.Add("3516", "亞帝歐");
            dics.Add("3520", "振維");
            dics.Add("3521", "鴻翊");
            dics.Add("3522", "御頂");
            dics.Add("3523", "迎輝");
            dics.Add("3526", "凡甲");
            dics.Add("3527", "聚積");
            dics.Add("3529", "力旺");
            dics.Add("3531", "先益");
            dics.Add("3537", "堡達");
            dics.Add("3540", "曜越");
            dics.Add("3541", "西柏");
            dics.Add("3546", "宇峻");
            dics.Add("3548", "兆利");
            dics.Add("3551", "世禾");
            dics.Add("3552", "同致");
            dics.Add("3555", "重鵬");
            dics.Add("3556", "禾瑞亞");
            dics.Add("3558", "神準");
            dics.Add("3562", "頂晶科");
            dics.Add("3563", "牧德");
            dics.Add("3564", "其陽");
            dics.Add("3567", "逸昌");
            dics.Add("3570", "大塚");
            dics.Add("3577", "泓格");
            dics.Add("3580", "友威科");
            dics.Add("3581", "博磊");
            dics.Add("3587", "閎康");
            dics.Add("3594", "磐儀");
            dics.Add("3609", "東林");
            dics.Add("3611", "鼎翰");
            dics.Add("3615", "安可");
            dics.Add("3623", "富晶通");
            dics.Add("3624", "光頡");
            dics.Add("3625", "西勝");
            dics.Add("3628", "盈正");
            dics.Add("3629", "地心引力");
            dics.Add("3630", "新鉅科");
            dics.Add("3631", "晟楠");
            dics.Add("3632", "研勤");
            dics.Add("3642", "駿熠電");
            dics.Add("3646", "艾恩特");
            dics.Add("3652", "精聯");
            dics.Add("3663", "鑫科");
            dics.Add("3664", "安瑞-KY");
            dics.Add("3666", "光耀");
            dics.Add("3672", "康聯訊");
            dics.Add("3675", "德微");
            dics.Add("3680", "家登");
            dics.Add("3684", "榮昌");
            dics.Add("3685", "元創精密");
            dics.Add("3687", "歐買尬");
            dics.Add("3689", "湧德");
            dics.Add("3691", "碩禾");
            dics.Add("3693", "營邦");
            dics.Add("3707", "漢磊");
            dics.Add("3709", "鑫聯大投控");
            dics.Add("3710", "連展投控");
            dics.Add("4102", "永日");
            dics.Add("4105", "東洋");
            dics.Add("4107", "邦特");
            dics.Add("4109", "穆拉德加捷");
            dics.Add("4111", "濟生");
            dics.Add("4113", "聯上");
            dics.Add("4114", "健喬");
            dics.Add("4116", "明基醫");
            dics.Add("4120", "友華");
            dics.Add("4121", "優盛");
            dics.Add("4123", "晟德");
            dics.Add("4126", "太醫");
            dics.Add("4127", "天良");
            dics.Add("4128", "中天");
            dics.Add("4129", "聯合");
            dics.Add("4130", "健亞");
            dics.Add("4131", "晶宇");
            dics.Add("4138", "曜亞");
            dics.Add("4139", "馬光-KY");
            dics.Add("4147", "中裕");
            dics.Add("4152", "台微體");
            dics.Add("4153", "鈺緯");
            dics.Add("4154", "康樂-KY");
            dics.Add("4157", "太景*-KY");
            dics.Add("4160", "創源");
            dics.Add("4161", "聿新科");
            dics.Add("4162", "智擎");
            dics.Add("4163", "鐿鈦");
            dics.Add("4167", "展旺");
            dics.Add("4168", "醣聯");
            dics.Add("4171", "瑞基");
            dics.Add("4173", "久裕");
            dics.Add("4174", "浩鼎");
            dics.Add("4175", "杏一");
            dics.Add("4180", "安成藥");
            dics.Add("4183", "福永生技");
            dics.Add("4188", "安克");
            dics.Add("4192", "杏國");
            dics.Add("4198", "環瑞醫");
            dics.Add("4205", "中華食");
            dics.Add("4207", "環泰");
            dics.Add("4303", "信立");
            dics.Add("4304", "勝昱");
            dics.Add("4305", "世坤");
            dics.Add("4401", "東隆興");
            dics.Add("4402", "福大");
            dics.Add("4406", "新昕纖");
            dics.Add("4413", "飛寶");
            dics.Add("4415", "台原藥");
            dics.Add("4416", "三圓");
            dics.Add("4417", "金洲");
            dics.Add("4419", "元勝");
            dics.Add("4420", "光明");
            dics.Add("4429", "聚紡");
            dics.Add("4430", "耀億");
            dics.Add("4432", "銘旺實");
            dics.Add("4433", "興采");
            dics.Add("4502", "健信");
            dics.Add("4503", "金雨");
            dics.Add("4506", "崇友");
            dics.Add("4510", "高鋒");
            dics.Add("4513", "福裕");
            dics.Add("4523", "永彰");
            dics.Add("4527", "方土霖");
            dics.Add("4528", "江興鍛");
            dics.Add("4529", "淳紳");
            dics.Add("4530", "宏易");
            dics.Add("4533", "協易機");
            dics.Add("4534", "慶騰");
            dics.Add("4535", "至興");
            dics.Add("4538", "大詠城");
            dics.Add("4541", "晟田");
            dics.Add("4542", "科嶠");
            dics.Add("4543", "萬在");
            dics.Add("4549", "桓達");
            dics.Add("4550", "長佳");
            dics.Add("4554", "橙的");
            dics.Add("4556", "旭然");
            dics.Add("4561", "健椿");
            dics.Add("4563", "百德");
            dics.Add("4568", "科際精密");
            dics.Add("4609", "唐鋒");
            dics.Add("4702", "中美實");
            dics.Add("4706", "大恭");
            dics.Add("4707", "磐亞");
            dics.Add("4711", "永純");
            dics.Add("4712", "南璋");
            dics.Add("4714", "永捷");
            dics.Add("4716", "大立");
            dics.Add("4721", "美琪瑪");
            dics.Add("4726", "永昕");
            dics.Add("4728", "雙美");
            dics.Add("4729", "熒茂");
            dics.Add("4735", "豪展");
            dics.Add("4736", "泰博");
            dics.Add("4741", "泓瀚");
            dics.Add("4743", "合一");
            dics.Add("4744", "皇將");
            dics.Add("4745", "合富-KY");
            dics.Add("4747", "強生");
            dics.Add("4754", "國碳科");
            dics.Add("4767", "誠泰科技");
            dics.Add("4803", "VHQ-KY");
            dics.Add("4804", "大略-KY");
            dics.Add("4806", "昇華");
            dics.Add("4903", "聯光通");
            dics.Add("4905", "台聯電");
            dics.Add("4907", "富宇");
            dics.Add("4908", "前鼎");
            dics.Add("4909", "新復興");
            dics.Add("4911", "德英");
            dics.Add("4924", "欣厚-KY");
            dics.Add("4933", "友輝");
            dics.Add("4939", "亞電");
            dics.Add("4944", "兆遠");
            dics.Add("4946", "辣椒");
            dics.Add("4947", "昂寶-KY");
            dics.Add("4950", "牧東");
            dics.Add("4953", "緯軟");
            dics.Add("4966", "譜瑞-KY");
            dics.Add("4971", "IET-KY");
            dics.Add("4972", "湯石照明");
            dics.Add("4973", "廣穎");
            dics.Add("4974", "亞泰");
            dics.Add("4979", "華星光");
            dics.Add("4987", "科誠");
            dics.Add("4991", "環宇-KY");
            dics.Add("4995", "晶達");
            dics.Add("5009", "榮剛");
            dics.Add("5011", "久陽");
            dics.Add("5013", "強新");
            dics.Add("5014", "建錩");
            dics.Add("5015", "華祺");
            dics.Add("5016", "松和");
            dics.Add("5102", "富強");
            dics.Add("5201", "凱衛");
            dics.Add("5202", "力新");
            dics.Add("5205", "中茂");
            dics.Add("5206", "坤悅");
            dics.Add("5209", "新鼎");
            dics.Add("5210", "寶碩");
            dics.Add("5211", "蒙恬");
            dics.Add("5212", "凌網");
            dics.Add("5213", "亞昕");
            dics.Add("5220", "萬達光電");
            dics.Add("5223", "安力-KY");
            dics.Add("5227", "立凱-KY");
            dics.Add("5230", "雷笛克光學");
            dics.Add("5245", "智晶");
            dics.Add("5251", "天鉞電");
            dics.Add("5263", "智崴");
            dics.Add("5272", "笙科");
            dics.Add("5274", "信驊");
            dics.Add("5276", "達輝-KY");
            dics.Add("5278", "尚凡");
            dics.Add("5281", "大峽谷-KY");
            dics.Add("5287", "數字");
            dics.Add("5289", "宜鼎");
            dics.Add("5291", "邑昇");
            dics.Add("5299", "杰力");
            dics.Add("5301", "寶得利");
            dics.Add("5302", "太欣");
            dics.Add("5304", "鼎創達");
            dics.Add("5306", "桂盟");
            dics.Add("5309", "系統電");
            dics.Add("5310", "天剛");
            dics.Add("5312", "寶島科");
            dics.Add("5314", "世紀");
            dics.Add("5315", "光聯");
            dics.Add("5317", "凱美");
            dics.Add("5321", "友銓");
            dics.Add("5324", "士開");
            dics.Add("5328", "華容");
            dics.Add("5340", "建榮");
            dics.Add("5344", "立衛");
            dics.Add("5345", "天揚");
            dics.Add("5347", "世界");
            dics.Add("5348", "系通");
            dics.Add("5349", "先豐");
            dics.Add("5351", "鈺創");
            dics.Add("5353", "台林");
            dics.Add("5355", "佳總");
            dics.Add("5356", "協益");
            dics.Add("5364", "力麗店");
            dics.Add("5371", "中光電");
            dics.Add("5381", "合正");
            dics.Add("5383", "金利");
            dics.Add("5386", "青雲");
            dics.Add("5392", "應華");
            dics.Add("5398", "慕康生醫");
            dics.Add("5403", "中菲");
            dics.Add("5410", "國眾");
            dics.Add("5425", "台半");
            dics.Add("5426", "振發");
            dics.Add("5432", "達威");
            dics.Add("5438", "東友");
            dics.Add("5439", "高技");
            dics.Add("5443", "均豪");
            dics.Add("5450", "寶聯通");
            dics.Add("5452", "佶優");
            dics.Add("5455", "昇益");
            dics.Add("5457", "宣德");
            dics.Add("5460", "同協");
            dics.Add("5464", "霖宏");
            dics.Add("5465", "富驊");
            dics.Add("5468", "凱鈺");
            dics.Add("5474", "聰泰");
            dics.Add("5475", "德宏");
            dics.Add("5478", "智冠");
            dics.Add("5480", "統盟");
            dics.Add("5481", "新華");
            dics.Add("5483", "中美晶");
            dics.Add("5487", "通泰");
            dics.Add("5488", "松普");
            dics.Add("5489", "彩富");
            dics.Add("5490", "同亨");
            dics.Add("5493", "三聯");
            dics.Add("5498", "凱崴");
            dics.Add("5508", "永信建");
            dics.Add("5511", "德昌");
            dics.Add("5512", "力麒");
            dics.Add("5514", "三豐");
            dics.Add("5516", "雙喜");
            dics.Add("5520", "力泰");
            dics.Add("5523", "豐謙");
            dics.Add("5529", "志嘉");
            dics.Add("5530", "龍巖");
            dics.Add("5536", "聖暉");
            dics.Add("5543", "崇佑-KY");
            dics.Add("5601", "台聯櫃");
            dics.Add("5603", "陸海");
            dics.Add("5604", "中連貨");
            dics.Add("5609", "中菲行");
            dics.Add("5701", "劍湖山");
            dics.Add("5703", "亞都");
            dics.Add("5704", "老爺知");
            dics.Add("5820", "日盛金");
            dics.Add("5864", "致和證");
            dics.Add("5878", "台名");
            dics.Add("5902", "德記");
            dics.Add("5903", "全家");
            dics.Add("5904", "寶雅");
            dics.Add("5905", "南仁湖");
            dics.Add("6015", "宏遠證");
            dics.Add("6016", "康和證");
            dics.Add("6020", "大展證");
            dics.Add("6021", "大慶證");
            dics.Add("6023", "元大期");
            dics.Add("6026", "福邦證");
            dics.Add("6101", "寬魚國際");
            dics.Add("6103", "合邦");
            dics.Add("6104", "創惟");
            dics.Add("6109", "亞元");
            dics.Add("6111", "大宇資");
            dics.Add("6113", "亞矽");
            dics.Add("6114", "久威");
            dics.Add("6118", "建達");
            dics.Add("6121", "新普");
            dics.Add("6122", "擎邦");
            dics.Add("6123", "上奇");
            dics.Add("6124", "業強");
            dics.Add("6125", "廣運");
            dics.Add("6126", "信音");
            dics.Add("6127", "九豪");
            dics.Add("6129", "普誠");
            dics.Add("6130", "星寶國際");
            dics.Add("6134", "萬旭");
            dics.Add("6138", "茂達");
            dics.Add("6140", "訊達");
            dics.Add("6143", "振曜");
            dics.Add("6144", "得利影");
            dics.Add("6146", "耕興");
            dics.Add("6147", "頎邦");
            dics.Add("6148", "驊宏資");
            dics.Add("6150", "撼訊");
            dics.Add("6151", "晉倫");
            dics.Add("6154", "順發");
            dics.Add("6156", "松上");
            dics.Add("6158", "禾昌");
            dics.Add("6160", "欣技");
            dics.Add("6161", "捷波");
            dics.Add("6163", "華電網");
            dics.Add("6167", "久正");
            dics.Add("6169", "昱泉");
            dics.Add("6170", "統振");
            dics.Add("6171", "亞銳士");
            dics.Add("6173", "信昌電");
            dics.Add("6174", "安碁");
            dics.Add("6175", "立敦");
            dics.Add("6179", "亞通");
            dics.Add("6180", "橘子");
            dics.Add("6182", "合晶");
            dics.Add("6185", "幃翔");
            dics.Add("6186", "新潤");
            dics.Add("6187", "萬潤");
            dics.Add("6188", "廣明");
            dics.Add("6190", "萬泰科");
            dics.Add("6194", "育富");
            dics.Add("6195", "詩肯");
            dics.Add("6198", "凌泰");
            dics.Add("6199", "天品");
            dics.Add("6203", "海韻電");
            dics.Add("6204", "艾華");
            dics.Add("6207", "雷科");
            dics.Add("6208", "日揚");
            dics.Add("6210", "慶生");
            dics.Add("6212", "理銘");
            dics.Add("6217", "中探針");
            dics.Add("6218", "豪勉");
            dics.Add("6219", "富旺");
            dics.Add("6220", "岳豐");
            dics.Add("6221", "晉泰");
            dics.Add("6222", "上揚");
            dics.Add("6223", "旺矽");
            dics.Add("6227", "茂綸");
            dics.Add("6228", "全譜");
            dics.Add("6229", "研通");
            dics.Add("6231", "系微");
            dics.Add("6233", "旺玖");
            dics.Add("6234", "高僑");
            dics.Add("6236", "康呈");
            dics.Add("6237", "驊訊");
            dics.Add("6238", "勝麗");
            dics.Add("6240", "松崗");
            dics.Add("6241", "易通展");
            dics.Add("6242", "立康");
            dics.Add("6244", "茂迪");
            dics.Add("6245", "立端");
            dics.Add("6246", "臺龍");
            dics.Add("6247", "淇譽電");
            dics.Add("6248", "沛波");
            dics.Add("6259", "百徽");
            dics.Add("6261", "久元");
            dics.Add("6263", "普萊德");
            dics.Add("6264", "富裔");
            dics.Add("6265", "方土昶");
            dics.Add("6266", "泰詠");
            dics.Add("6270", "倍微");
            dics.Add("6274", "台燿");
            dics.Add("6275", "元山");
            dics.Add("6276", "安鈦克");
            dics.Add("6279", "胡連");
            dics.Add("6284", "佳邦");
            dics.Add("6287", "元隆");
            dics.Add("6290", "良維");
            dics.Add("6291", "沛亨");
            dics.Add("6292", "迅德");
            dics.Add("6294", "智基");
            dics.Add("6404", "通訊-KY");
            dics.Add("6411", "晶焱");
            dics.Add("6417", "韋僑");
            dics.Add("6418", "詠昇");
            dics.Add("6419", "京晨科");
            dics.Add("6425", "易發");
            dics.Add("6426", "統新");
            dics.Add("6432", "今展科");
            dics.Add("6435", "大中");
            dics.Add("6438", "迅得");
            dics.Add("6441", "廣錠");
            dics.Add("6446", "藥華藥");
            dics.Add("6457", "紘康");
            dics.Add("6461", "益得");
            dics.Add("6462", "神盾");
            dics.Add("6465", "威潤");
            dics.Add("6469", "大樹");
            dics.Add("6470", "宇智");
            dics.Add("6472", "保瑞");
            dics.Add("6482", "弘煜科");
            dics.Add("6485", "點序");
            dics.Add("6486", "互動");
            dics.Add("6488", "環球晶");
            dics.Add("6492", "生華科");
            dics.Add("6494", "九齊");
            dics.Add("6496", "科懋");
            dics.Add("6497 ", "亞獅康-KY");
            dics.Add("6499", "益安");
            dics.Add("6506", "雙邦");
            dics.Add("6508", "惠光");
            dics.Add("6509", "聚和");
            dics.Add("6510", "精測");
            dics.Add("6512", "啟發電");
            dics.Add("6514", "芮特-KY");
            dics.Add("6523", "達爾膚");
            dics.Add("6530", "創威");
            dics.Add("6532", "瑞耘");
            dics.Add("6535", "順藥");
            dics.Add("6538", "倉和");
            dics.Add("6542", "隆中");
            dics.Add("6547", "高端疫苗");
            dics.Add("6548", "長科");
            dics.Add("6556", "勝品");
            dics.Add("6560", "欣普羅");
            dics.Add("6561", "是方");
            dics.Add("6568", "宏觀");
            dics.Add("6569", "醫揚");
            dics.Add("6570", "維田");
            dics.Add("6574", "霈方");
            dics.Add("6576", "逸達");
            dics.Add("6577", "勁豐");
            dics.Add("6578", "達邦蛋白");
            dics.Add("6590", "普鴻");
            dics.Add("6593", "台灣銘板");
            dics.Add("6594", "展匯科");
            dics.Add("6596", "寬宏藝術");
            dics.Add("6603", "富強鑫");
            dics.Add("6609", "瀧澤科");
            dics.Add("6612", "奈米醫材");
            dics.Add("6613", "朋億");
            dics.Add("6615", "慧智");
            dics.Add("6616", "特昇-KY");
            dics.Add("6640", "均華");
            dics.Add("6643", "M31");
            dics.Add("6654", "天正國際");
            dics.Add("6664", "群翊");
            dics.Add("6667", "信紘科");
            dics.Add("6803", "崑鼎");
            dics.Add("7402", "邑錡");
            dics.Add("8024", "佑華");
            dics.Add("8027", "鈦昇");
            dics.Add("8032", "光菱");
            dics.Add("8034", "榮群");
            dics.Add("8038", "長園科");
            dics.Add("8040", "九暘");
            dics.Add("8042", "金山電");
            dics.Add("8043", "蜜望實");
            dics.Add("8044", "網家");
            dics.Add("8047", "星雲");
            dics.Add("8048", "德勝");
            dics.Add("8049", "晶采");
            dics.Add("8050", "廣積");
            dics.Add("8054", "安國");
            dics.Add("8059", "凱碩");
            dics.Add("8064", "東捷");
            dics.Add("8066", "來思達");
            dics.Add("8067", "志旭");
            dics.Add("8068", "全達");
            dics.Add("8069", "元太");
            dics.Add("8071", "能率網通");
            dics.Add("8074", "鉅橡");
            dics.Add("8076", "伍豐");
            dics.Add("8077", "洛碁");
            dics.Add("8080", "奧斯特");
            dics.Add("8083", "瑞穎");
            dics.Add("8084", "巨虹");
            dics.Add("8085", "福華");
            dics.Add("8086", "宏捷科");
            dics.Add("8087", "華鎂鑫");
            dics.Add("8088", "品安");
            dics.Add("8091", "翔名");
            dics.Add("8092", "建暐");
            dics.Add("8093", "保銳");
            dics.Add("8096", "擎亞");
            dics.Add("8097", "常珵");
            dics.Add("8099", "大世科");
            dics.Add("8107", "大億金茂");
            dics.Add("8109", "博大");
            dics.Add("8111", "立碁");
            dics.Add("8121", "越峰");
            dics.Add("8147", "正淩");
            dics.Add("8155", "博智");
            dics.Add("8171", "天宇");
            dics.Add("8176", "智捷");
            dics.Add("8182", "加高");
            dics.Add("8183", "精星");
            dics.Add("8234", "新漢");
            dics.Add("8240", "華宏");
            dics.Add("8255", "朋程");
            dics.Add("8277", "商丞");
            dics.Add("8279", "生展");
            dics.Add("8287", "英格爾");
            dics.Add("8289", "泰藝");
            dics.Add("8291", "尚茂");
            dics.Add("8299", "群聯");
            dics.Add("8342", "益張");
            dics.Add("8349", "恒耀");
            dics.Add("8354", "冠好");
            dics.Add("8358", "金居");
            dics.Add("8383", "千附");
            dics.Add("8390", "金益鼎");
            dics.Add("8401", "白紗科");
            dics.Add("8403", "盛弘");
            dics.Add("8406", "金可-KY");
            dics.Add("8409", "商之器");
            dics.Add("8410", "森田");
            dics.Add("8415", "大國鋼");
            dics.Add("8416", "實威");
            dics.Add("8418", "捷必勝-KY");
            dics.Add("8420", "明揚");
            dics.Add("8421", "旭源");
            dics.Add("8423", "保綠-KY");
            dics.Add("8424", "惠普");
            dics.Add("8426", "紅木-KY");
            dics.Add("8431", "匯鑽科");
            dics.Add("8432", "東生華");
            dics.Add("8433", "弘帆");
            dics.Add("8435", "鉅邁");
            dics.Add("8436", "大江");
            dics.Add("8437", "大地-KY");
            dics.Add("8440", "綠電");
            dics.Add("8444", "綠河-KY");
            dics.Add("8446", "華研");
            dics.Add("8450", "霹靂");
            dics.Add("8455", "大拓-KY");
            dics.Add("8462", "柏文");
            dics.Add("8472", "夠麻吉");
            dics.Add("8476", "台境");
            dics.Add("8477", "創業家");
            dics.Add("8489", "三貝德");
            dics.Add("8905", "裕國");
            dics.Add("8906", "花王");
            dics.Add("8908", "欣雄");
            dics.Add("8913", "全銓");
            dics.Add("8916", "光隆");
            dics.Add("8917", "欣泰");
            dics.Add("8921", "沈氏");
            dics.Add("8923", "時報");
            dics.Add("8924", "大田");
            dics.Add("8927", "北基");
            dics.Add("8928", "鉅明");
            dics.Add("8929", "富堡");
            dics.Add("8930", "青鋼");
            dics.Add("8931", "大汽電");
            dics.Add("8932", "宏大");
            dics.Add("8933", "愛地雅");
            dics.Add("8934", "衡平");
            dics.Add("8935", "邦泰");
            dics.Add("8936", "國統");
            dics.Add("8937", "合騏");
            dics.Add("8938", "明安");
            dics.Add("8941", "關中");
            dics.Add("8942", "森鉅");
            dics.Add("9949", "琉園");
            dics.Add("9950", "萬國通");
            dics.Add("9951", "皇田");
            dics.Add("9960", "邁達康");
            dics.Add("9962", "有益");


            return dics;
        }

        private void chkThread() {
            if (thread != null)
            {
                if (thread.ThreadState != ThreadState.Aborted)
                {
                    thread.Abort();
                    //檢測執行續是否停止
                  

                    thread = null;
                }
            }

        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            chkThread();

            IsTrans = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            chkThread();

            Application.ExitThread();
            Environment.Exit(0);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            //if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday || DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
            //{
            //    return;
            //}

            //if (DateTime.Now.Hour == 15)
            //{
            //    btnTransfer_Click(null, null);
            //}
        }
    }
}
