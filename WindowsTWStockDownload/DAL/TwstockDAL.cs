using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Dapper;
namespace WindowsTWStockDownload.DAL
{
    /// <summary>[twstock]表数据访问类
    /// 作者:alonso(line:menshin7)
    /// 创建时间:2019-04-04 22:45:49
    /// </summary>
    public partial class TwstockDAL
    {
        public TwstockDAL()
        { }
        /// <summary>增加一条数据
        /// 
        /// </summary>
        public int Add(WindowsTWStockDownload.Model.Twstock model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into [twstock](");
            strSql.Append("[chDate], [ClosePrice], [UpOrDown], [TradeCount], [TradeMoney], [OpenPrice], [HightPrice], [LowPrice], [StockCode], [CreateDate]  )");
            strSql.Append(" values (");
            strSql.Append("@chDate, @ClosePrice, @UpOrDown, @TradeCount, @TradeMoney, @OpenPrice, @HightPrice, @LowPrice, @StockCode, @CreateDate  );select @@identity;");

            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                int i = connection.Query<int>(strSql.ToString(), model).SingleOrDefault(); ;
                return i;
            }
        }

        /// <summary>更新一条数据
        /// 
        /// </summary>
        public bool Update(WindowsTWStockDownload.Model.Twstock model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update [twstock] set ");
            strSql.Append("[chDate]=@chDate, [ClosePrice]=@ClosePrice, [UpOrDown]=@UpOrDown, [TradeCount]=@TradeCount, [TradeMoney]=@TradeMoney, [OpenPrice]=@OpenPrice, [HightPrice]=@HightPrice, [LowPrice]=@LowPrice, [StockCode]=@StockCode, [CreateDate]=@CreateDate  ");
            strSql.Append(" where Id=@Id ");
            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                int i = connection.Execute(strSql.ToString(), model);
                return i > 0;
            }
        }

        /// <summary>按条件更新数据
        /// 
        /// </summary>
        public bool UpdateByCond(string str_set, string cond)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update [twstock] set " + str_set + " ");
            strSql.Append(" where " + cond);
            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                int i = connection.Execute(strSql.ToString());
                return i > 0;
            }
        }

        /// <summary>删除一条数据
        /// 
        /// </summary>
        public bool Delete(Guid Id)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from [twstock] ");
            strSql.Append(" where Id=@Id ");
            using (var connection = ConnectionFactory.GetOpenConnection())
            {

                int i = connection.Execute(strSql.ToString(), new { Id = Id });
                return i > 0;
            }
        }

        /// <summary>根据条件删除数据
        /// 
        /// </summary>
        public bool DeleteByCond(string cond)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from [twstock] ");
            if (!string.IsNullOrEmpty(cond))
            {
                strSql.Append(" where " + cond);
            }
            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                int i = connection.Execute(strSql.ToString());
                return i > 0;
            }
        }

        /// <summary>取一个字段的值
        /// 
        /// </summary>
        /// <param name="filed">字段，如sum(je)</param>
        /// <param name="cond">条件，如userid=2</param>
        /// <returns></returns>
        public string GetOneFiled(string filed, string cond)
        {
            string sql = "select " + filed + " from [twstock]";
            if (!string.IsNullOrEmpty(cond))
            {
                sql += " where " + cond;
            }

            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                string tmp = connection.ExecuteScalar<string>(sql);
                return tmp;
            }
        }

        /// <summary>取一个字段的值
        /// 
        /// </summary>
        /// <param name="filed">字段，如sum(je)</param>
        /// <param name="cond">条件，如userid=2</param>
        /// <returns></returns>
        public double GetOneFiledDouble(string filed, string cond)
        {
            string tmp = GetOneFiled(filed, cond);
            return string.IsNullOrEmpty(tmp) ? 0 : double.Parse(tmp);
        }

        /// <summary>得到一个对象实体
        /// 
        /// </summary>
        public WindowsTWStockDownload.Model.Twstock GetModel(Guid Id)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select * from [twstock] ");
            strSql.Append(" where Id=@Id ");
            WindowsTWStockDownload.Model.Twstock model = null;
            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                model = connection.Query<WindowsTWStockDownload.Model.Twstock>(strSql.ToString(), new { Id = Id }).SingleOrDefault();

            }
            return model;
        }

        /// <summary>根据条件得到一个对象实体
        /// 
        /// </summary>
        public WindowsTWStockDownload.Model.Twstock GetModelByCond(string cond)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select top 1 * from [twstock] ");
            if (!string.IsNullOrEmpty(cond))
            {
                strSql.Append(" where " + cond);
            }
            WindowsTWStockDownload.Model.Twstock model = null;
            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                model = connection.Query<WindowsTWStockDownload.Model.Twstock>(strSql.ToString()).SingleOrDefault();

            }
            return model;
        }






        /// <summary>获得数据列表
        /// 
        /// </summary>
        public List<WindowsTWStockDownload.Model.Twstock> GetListArray(string strWhere)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select * ");
            strSql.Append(" FROM [twstock] ");
            if (strWhere.Trim() != "")
            {
                strSql.Append(" where " + strWhere);
            }
            List<WindowsTWStockDownload.Model.Twstock> list = new List<WindowsTWStockDownload.Model.Twstock>();
            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                list = connection.Query<WindowsTWStockDownload.Model.Twstock>(strSql.ToString()).ToList();

            }
            return list;
        }

        /// <summary>分页获取数据列表
        /// 
        /// </summary>
        public List<WindowsTWStockDownload.Model.Twstock> GetListArray(string fileds, string orderstr, int PageSize, int PageIndex, string strWhere)
        {
            string order = orderstr.Split(' ')[0];
            string ordertype = orderstr.Split(' ')[1];
            string cond = string.IsNullOrEmpty(strWhere) ? "" : string.Format(" where {0}", strWhere);
            string sql = string.Format("SELECT * FROM ( SELECT ROW_NUMBER() OVER (ORDER BY {0} {1}) AS pos, {2} FROM  [twstock] {3}  ) AS sp WHERE pos BETWEEN {4} AND {5}", order, ordertype, fileds, cond, (((PageIndex - 1) * PageSize) + 1), PageSize * PageIndex);

            // 	    string sql = string.Format("select {0} from [twstock] {1} order by {2} offset {3} rows fetch next {4} rows only", fileds, cond, orderstr, (PageIndex - 1) * PageSize, PageSize);


            List<WindowsTWStockDownload.Model.Twstock> list = new List<WindowsTWStockDownload.Model.Twstock>();
            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                list = connection.Query<WindowsTWStockDownload.Model.Twstock>(sql).ToList();

            }
            return list;
        }



        /// <summary>计算记录数
        /// 
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public int CalcCount(string cond)
        {
            string sql = "select count(1) from [twstock]";
            if (!string.IsNullOrEmpty(cond))
            {
                sql += " where " + cond;
            }
            using (var connection = ConnectionFactory.GetOpenConnection())
            {
                int i = connection.Query<int>(sql).SingleOrDefault();
                return i;

            }
        }

        //1234

    }
}

