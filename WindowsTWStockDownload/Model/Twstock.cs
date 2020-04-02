using System;
namespace WindowsTWStockDownload.Model
{
    /// <summary>twstock表实体类
    /// 作者:alonso(line:menshin7)
    /// 创建时间:2019-04-04 22:46:49
    /// </summary>
    [Serializable]
    public partial class Twstock
    {
        public Twstock()
        { }
        private Guid _Id;
        /// <summary>
        /// 
        /// </summary>
        public Guid Id
        {
            set { _Id = value; }
            get { return _Id; }
        }
        private string _chDate;
        /// <summary>
        /// 
        /// </summary>
        public string chDate
        {
            set { _chDate = value; }
            get { return _chDate; }
        }
        private double? _ClosePrice;
        /// <summary>
        /// 
        /// </summary>
        public double? ClosePrice
        {
            set { _ClosePrice = value; }
            get { return _ClosePrice; }
        }
        private string _UpOrDown;
        /// <summary>
        /// 
        /// </summary>
        public string UpOrDown
        {
            set { _UpOrDown = value; }
            get { return _UpOrDown; }
        }
        private string _TradeCount;
        /// <summary>
        /// 
        /// </summary>
        public string TradeCount
        {
            set { _TradeCount = value; }
            get { return _TradeCount; }
        }
        private string _TradeMoney;
        /// <summary>
        /// 
        /// </summary>
        public string TradeMoney
        {
            set { _TradeMoney = value; }
            get { return _TradeMoney; }
        }
        private string _OpenPrice;
        /// <summary>
        /// 
        /// </summary>
        public string OpenPrice
        {
            set { _OpenPrice = value; }
            get { return _OpenPrice; }
        }
        private string _HightPrice;
        /// <summary>
        /// 
        /// </summary>
        public string HightPrice
        {
            set { _HightPrice = value; }
            get { return _HightPrice; }
        }
        private string _LowPrice;
        /// <summary>
        /// 
        /// </summary>
        public string LowPrice
        {
            set { _LowPrice = value; }
            get { return _LowPrice; }
        }
        private string _StockCode;
        /// <summary>
        /// 
        /// </summary>
        public string StockCode
        {
            set { _StockCode = value; }
            get { return _StockCode; }
        }
        private DateTime? _CreateDate = DateTime.Now;
        /// <summary>
        /// 
        /// </summary>
        public DateTime? CreateDate
        {
            set { _CreateDate = value; }
            get { return _CreateDate; }
        }
    }
}
