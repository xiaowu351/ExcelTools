using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.OleDb;
using System.IO;

namespace CatchExcel
{
    public partial class frmCatchExcel : Form
    {
        #region 自定义变量
        private const string ISYES = "是";
        private const string ISNO = "否";
        private const string BHZC1 = "依安全";
        private const string BHZC2 = "依需求";
        private const string BHZC3 = "不需補";
        private const string ZYLY1 = "生產件";
        /// <summary>
        /// 廠內生產件 2014.12.23增加
        /// </summary>
        private const string ZYLY11 = "廠內生產件";
        /// <summary>
        /// 采购件 ：2014.12.23 输出修改为3。原来输出为2
        /// </summary>
        private const string ZYLY2_3 = "採購件";
        private const string ZYLY3 = "托工件";
        private const string ZDGL1 = "逐批領料";
        private const string ZDGL2 = "自動扣料";
        private const string JZFS1 = "現結";
        private const string JZFS2 = "月結";
        #endregion
        #region Access查詢語句
        /// <summary>
        /// 商品
        /// </summary>
        private static string AccessSqlProduct_imscj = @"select 品號,品名,規格,
			 '' as 固定1,'' as 固定2,'' as 固定3,
			庫存單位,'' as 固定5,'False' as 布尔1,'True' as 布尔2,
			'' as 固定6,'' as 固定7,循環盤點碼,
			'' as 固定8,'' as 固定9,'' as 固定10,'' as 固定11,
			'Z:\111\{0}.jpg' as 图片路径,'' as 固定12,
			主供應商,補貨政策,'0' as 固定13,安全存量,
			'' as 固定14,'' as 固定15,補貨倍量,標準進價,'False' as 布尔3,
			'' as 固定16,'' as 固定17,幣別,'' as 固定18,
			零售價,'False' as 布尔4,
			'' as 固定19,'' as 固定20,'' as 固定21,'' as 固定22,
			'' as 固定23,'' as 固定24,'' as 固定25,'' as 固定26,
			'' as 固定27,'' as 固定28,'' as 固定29,'' as 固定30,
			'' as 固定31,'' as 固定32,'' as 固定33,'' as 固定34,
			儲位,'' as 固定35,'' as 固定36,'' as 固定37,'' as 固定38,
			'' as 固定39,'' as 固定40,'False' as 布尔5,
			'' as 固定41,'' as 固定42,'' as 固定43,'' as 固定44,
			'' as 固定45,'' as 固定46,主要來源,主要庫別,'' as 固定47,'1' as 固定48,
			'' as 固定49,'' as 固定50,'' as 固定51,'' as 固定52,
			'' as 固定53,'' as 固定54,'' as 固定55,'' as 固定56,
			'' as 固定57,'' as 固定58,'' as 固定59,'' as 固定60,
			'' as 固定61,'DS' as 固定62,'' as 固定63,'DS' as 固定64,
			'' as 固定65,'' as 固定66,計劃人員 as 固定67,'' as 固定68,
			'' as 固定69,'' as 固定70,'' as 固定71,'End' from [Sheet1$]  where 品號 {1}''";
        /// <summary>
        /// 库存
        /// </summary>
        private static string AccessSqlProduct_imsco = @"select 'AA' as 固定1,'imsco{0}',品號,品名,規格,庫存單位,'0' as 固定2,主要庫別,庫存數量, '' as 固定3,'' as 固定4,'' as 固定5,'' as 固定6, '' as 固定7,'' as 固定8,'' as 固定9,'' as 固定10,'' as 固定11,'True' as 布尔1,'' as 固定12,'' as 固定13,'DS' as 固定14,'' as 固定15,'DS' as 固定16,'' as 固定17,'' as 固定18,'End' from [Sheet1$]  where 品號 {1}''";
        /// <summary>
        /// 厂商
        /// </summary>
        private static string AccessSqlProduct_imsdb = @"select 廠商代號,廠商全稱,廠商簡稱,'' as 固定1,'' as 固定2,
負責人,聯系人,電話號碼一,'' as 固定3,傳真,
Email,廠商地址,'' as 固定4,'' as 固定5,'' as 固定6,'' as 固定7,
交易幣別,'2' as 固定8,'1' as 固定9,'' as 固定10,結帳方式,
'' as 固定11,'' as 固定12,'' as 固定13,'' as 固定14,
'' as 固定15,'' as 固定16,'' as 固定17,'DS' as 固定18,'' as 固定19,
'DS' as 固定20,'' as 固定21,'' as 固定22,'' as 固定23,聯系人,'' as 固定24,
廠商地址,'' as 固定25,'' as 固定26,'' as 固定27,'' as 固定28,'' as 固定29,'' as 固定30,
'False' as 布尔1,
'' as 固定31,'' as 固定32,'' as 固定33,'' as 固定34,
'' as 固定35,'' as 固定36,'' as 固定37,'' as 固定38,
'' as 固定39,'' as 固定40,'' as 固定41,'' as 固定42,
'' as 固定43,'' as 固定44,'' as 固定45,'' as 固定46,
'' as 固定47,'' as 固定48,'' as 固定49, '0' as 固定50,
'' as 固定51,'' as 固定52,'' as 固定53,'' as 固定54,'' as 固定55,'' as 固定56,'End' from [Sheet1$]  where 廠商代號 {0}'' ";
        /// <summary>
        /// 客户
        /// </summary>
        private static string AccessSqlProduct_imsef = @"select 客戶代號,客戶全名,客戶簡稱,
'' as 固定1,負責人,聯系人,
'' as 固定2,'' as 固定3,'' as 固定4,'' as 固定5,
'' as 固定6,'' as 固定7,'' as 固定8,'' as 固定9,
'' as 固定10,'' as 固定11, 電話號碼一,'' as 固定88,傳真,Email,
發票地址,發票地址,發票地址,
'' as 固定12,'' as 固定13,'' as 固定14,'' as 固定15,
'' as 固定16,'' as 固定17,'' as 固定18,'2' as 固定19,
'1' as 固定20,交易幣別,
'1' as 固定21,'1' as 固定22,'' as 固定23,'' as 固定24,
結帳方式,'' as 固定26,'' as 固定27,'' as 固定28,
'' as 固定29,'' as 固定30,'' as 固定31,'' as 固定32,
'' as 固定33,'' as 固定34,客戶代號,
'' as 固定35,'' as 固定36,'' as 固定37,'' as 固定38,
'' as 固定39,'' as 固定40,'' as 固定41,'' as 固定42,
'' as 固定43,'' as 固定44,'' as 固定45,'' as 固定46,
'' as 固定47,'' as 固定48,'' as 固定49,'' as 固定50,
'' as 固定51,'' as 固定52,'' as 固定53,'' as 固定54,
'' as 固定55,'' as 固定56,'' as 固定57,'' as 固定58,
'' as 固定59,'' as 固定60,'' as 固定61,'' as 固定62,
'' as 固定63,'' as 固定64,'' as 固定65,'' as 固定66,
'' as 固定67,'' as 固定68,'' as 固定69,'' as 固定70,
'' as 固定71,'' as 固定72,'DS' as 固定73,
'' as 固定74,'' as 固定75,'' as 固定76,'' as 固定77,
'' as 固定78,'' as 固定79,'' as 固定80,'' as 固定81,
'' as 固定82,'' as 固定83,'' as 固定84,'' as 固定85,
'' as 固定86,'' as 固定87,'End' from [Sheet1$]  where 客戶代號 {0}''";
        #endregion
       

        /// <summary>
        /// 保存用户选择的文件名及路径
        /// </summary>
        private  string fileName = string.Empty;

        public frmCatchExcel()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 商品抓取
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnimscj_Click(object sender, EventArgs e)
        {
            // 1.获取打开的Excel文件 
            
            if (string.IsNullOrEmpty(this.GetSelectFilePathName()))
            {
                return;
            }

            //string sql = ConfigurationManager.AppSettings["AccessSqlProduct_imscj"];
            string sql = AccessSqlProduct_imscj;
            DataTable dt = ExcelToDataSet(fileName, string.Format(sql,"{0}","<>"));
            this.dgvProduct.DataSource = dt;
            string filePath = fileName.Substring(0, fileName.LastIndexOf("\\"));
            filePath += "\\商品IMSCJ" + new Random().Next() + ".txt";
            //写入到txt文件中
            WriteTxt(filePath,dt);
           
            
        }

        /// <summary>
        /// 将文件写入txt中
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="dt"></param>
        private void WriteTxt(string filePath, DataTable dt)
        {
           
            StringBuilder builder = new StringBuilder();
            if (dt == null || dt.Rows.Count <= 0)
            {
                MessageBox.Show("无数据可写入txt");
                return;
            }
            DateTime dateCell;
            Decimal decimalCell;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (Decimal.TryParse(dt.Rows[i][j].ToString(), out decimalCell))
                    {//数值类型
                        builder.AppendFormat("\"{0}\",", Convert.ToDecimal(dt.Rows[i][j]));
                    }
                    else  if (DateTime.TryParse(dt.Rows[i][j].ToString(), out dateCell))
                    {//日期类型
                        builder.AppendFormat("\"{0}\",", dateCell.ToString("yyyy-MM-dd").Replace("-", ""));
                    }
                   
                    else if (dt.Rows[i][j].ToString().Contains(".jpg"))
                    {
                        builder.AppendFormat("\"{0}\",", string.Format(dt.Rows[i][j].ToString(), dt.Rows[i][0]));
                    }
                    else if (dt.Rows[i][j].ToString()==ISYES)
                    {
                        builder.AppendFormat("\"{0}\",", "True");
                    }
                    else if (dt.Rows[i][j].ToString() == ISNO)
                    {
                        builder.AppendFormat("\"{0}\",", "False");
                    }
                    #region 補貨政策、主要來源、結帳方式
                    else if (dt.Rows[i][j].ToString() == BHZC1 || dt.Rows[i][j].ToString() == ZYLY1 || dt.Rows[i][j].ToString() == ZYLY11 || dt.Rows[i][j].ToString() == ZDGL1 || dt.Rows[i][j].ToString() == JZFS1)
                    {
                        builder.AppendFormat("\"{0}\",", "1");
                    }
                    else if (dt.Rows[i][j].ToString() == BHZC2 || dt.Rows[i][j].ToString() == ZDGL2 || dt.Rows[i][j].ToString() == JZFS2)
                    {
                        builder.AppendFormat("\"{0}\",", "2");
                    }
                    else if (dt.Rows[i][j].ToString() == BHZC3 || dt.Rows[i][j].ToString() == ZYLY3 || dt.Rows[i][j].ToString() == ZYLY2_3)
                    {
                        builder.AppendFormat("\"{0}\",", "3");
                    }
                    #endregion
                    else if (dt.Rows[i][j].ToString().Contains("End"))
                    {
                        builder.AppendFormat("\"\"<<{0}>>",dt.Rows[i][j]);
                    }
                    else if (dt.Rows[i][j].ToString().Contains("imsco"))
                    {

                        builder.AppendFormat("\"{0}\",", string.Format(dt.Rows[i][j].ToString().Replace("imsco", ""), (i + 1)).PadLeft(4, '0'));
                    } 
                    else  
                    {
                        builder.AppendFormat("\"{0}\",", dt.Rows[i][j]);
                    } 
                }
                builder.AppendFormat("\r\n");
            }
            try
            {
                File.WriteAllText(filePath, builder.ToString().TrimEnd(','), Encoding.Default);
                MessageBox.Show("抓取完畢！\r\n文件保存至："+filePath);
            }
            catch (Exception ex)
            {

                MessageBox.Show("發生錯誤："+ex.Message);
            }
            
        }


        /// <summary>
        /// 读取Excel文件，将内容存储在DataSet中
        /// </summary>
        /// <param name="opnFileName">带路径的Excel文件名</param>
        /// <returns>DataSet</returns>
        public static DataTable ExcelToDataSet(string opnFileName, string SqlExcel)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + opnFileName + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            OleDbConnection conn = new OleDbConnection(strConn);
            // string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = new DataSet();

             
            try
            {
                conn.Open();
                myCommand = new OleDbDataAdapter(SqlExcel, strConn);

                myCommand.Fill(ds, "dbSource");
                return ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("导入出错：\r\n" + ex.Message, "\n\r错误信息");
                return null;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
        }

        /// <summary>
        /// 获取文件名
        /// </summary>
        /// <returns></returns>
        public  string GetSelectFilePathName()
        {
            // 1.获取打开的Excel文件 
            this.openFileDialog1.Filter = "Excel文件|*.xls;*.xlsx;";
            if (this.openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return string.Empty;
            }
            this.fileName = openFileDialog1.FileName;
            return fileName;
        }

        /// <summary>
        /// 库存抓取
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImscl_Click(object sender, EventArgs e)
        {
            // 1.获取打开的Excel文件 

            if (string.IsNullOrEmpty(this.GetSelectFilePathName()))
            {
                return;
            }
           // string sql = ConfigurationManager.AppSettings["AccessSqlProduct_imscl"];
            string sql = AccessSqlProduct_imsco;
            DataTable dt = ExcelToDataSet(fileName, string.Format(sql, "{0}", "<>"));
            this.dgvProduct.DataSource = dt;
            //string filePath = fileName.Replace(".xls", DateTime.Now.Ticks + ".txt");
            string filePath = fileName.Substring(0, fileName.LastIndexOf("\\"));
            filePath +="\\庫存IMSCL"+new Random().Next()+ ".txt";
            //写入到txt文件中
            WriteTxt(filePath, dt);
        }

        /// <summary>
        /// 抓取厂商
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImsdb_Click(object sender, EventArgs e)
        {
            // 1.获取打开的Excel文件 

            if (string.IsNullOrEmpty(this.GetSelectFilePathName()))
            {
                return;
            }
            string sql = AccessSqlProduct_imsdb;
            DataTable dt = ExcelToDataSet(fileName, string.Format(sql,"<>"));
            this.dgvProduct.DataSource = dt;
            //string filePath = fileName.Replace(".xls", DateTime.Now.Ticks + ".txt");
            string filePath = fileName.Substring(0, fileName.LastIndexOf("\\"));
            filePath += "\\廠商IMSDB" + new Random().Next() + ".txt";
            //写入到txt文件中
            WriteTxt(filePath, dt);
        }

        /// <summary>
        /// 客戶抓取
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImsdf_Click(object sender, EventArgs e)
        {
            // 1.获取打开的Excel文件 

            if (string.IsNullOrEmpty(this.GetSelectFilePathName()))
            {
                return;
            }
            string sql = AccessSqlProduct_imsef;
            DataTable dt = ExcelToDataSet(fileName, string.Format(sql, "<>"));
            this.dgvProduct.DataSource = dt;
            //string filePath = fileName.Replace(".xls", DateTime.Now.Ticks + ".txt");
            string filePath = fileName.Substring(0, fileName.LastIndexOf("\\"));
            filePath += "\\客戶IMSEF" + new Random().Next() + ".txt";
            //写入到txt文件中
            WriteTxt(filePath, dt);
        }
        
    }
}
