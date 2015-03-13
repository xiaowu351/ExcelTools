using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CatchExcel
{
	/// <summary>
	/// 抓取工具帮助类
	/// </summary>
	public static class CatchExcelHelper
	{

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
				throw new Exception("分析出錯：\r\n" + ex.Message);
				//MessageBox.Show("导入出错：\r\n" + ex.Message, "\n\r错误信息");
				//return null;
			}
			finally
			{
				conn.Close();
				conn.Dispose();
			}
		}


		/// <summary>
		/// 将文件写入txt中
		/// </summary>
		/// <param name="filePath"></param>
		/// <param name="dt"></param>
		public static void WriteTxt(string filePath, DataTable dt)
		{

			StringBuilder builder = new StringBuilder();
			if (dt == null || dt.Rows.Count <= 0)
			{
				throw new Exception("無數據可寫入txt");
			}
			DateTime dateCell;
			DateTime doDataTime = DateTime.Now;
			Decimal decimalCell;
			for (int i = 0; i < dt.Rows.Count; i++)
			{
				for (int j = 0; j < dt.Columns.Count; j++)
				{
					if (Decimal.TryParse(dt.Rows[i][j].ToString(), out decimalCell))
					{//数值类型
						builder.AppendFormat("\"{0}\",", Convert.ToDecimal(dt.Rows[i][j]));
					}
					else if (DateTime.TryParse(dt.Rows[i][j].ToString(), out dateCell))
					{//日期类型
						builder.AppendFormat("\"{0}\",", dateCell.ToString("yyyy-MM-dd").Replace("-", ""));
					}

					else if (dt.Rows[i][j].ToString().Contains(".jpg"))
					{
						builder.AppendFormat("\"{0}\",", string.Format(dt.Rows[i][j].ToString(), dt.Rows[i][0]));
					}
					else if (dt.Rows[i][j].ToString() == ConstDefine.ISYES)
					{
						builder.AppendFormat("\"{0}\",", "True");
					}
					else if (dt.Rows[i][j].ToString() == ConstDefine.ISNO)
					{
						builder.AppendFormat("\"{0}\",", "False");
					}
					#region 補貨政策、主要來源、結帳方式
					else if (dt.Rows[i][j].ToString() == ConstDefine.BHZC1
							|| dt.Rows[i][j].ToString() == ConstDefine.ZYLY1
							|| dt.Rows[i][j].ToString() == ConstDefine.ZYLY11
							|| dt.Rows[i][j].ToString() == ConstDefine.ZDGL1
							|| dt.Rows[i][j].ToString() == ConstDefine.JZFS1)
					{
						builder.AppendFormat("\"{0}\",", "1");
					}
					else if (dt.Rows[i][j].ToString() == ConstDefine.BHZC2
							|| dt.Rows[i][j].ToString() == ConstDefine.ZDGL2
							|| dt.Rows[i][j].ToString() == ConstDefine.JZFS2)
					{
						builder.AppendFormat("\"{0}\",", "2");
					}
					else if (dt.Rows[i][j].ToString() == ConstDefine.BHZC3
						|| dt.Rows[i][j].ToString() == ConstDefine.ZYLY3
						|| dt.Rows[i][j].ToString() == ConstDefine.ZYLY2_3)
					{
						builder.AppendFormat("\"{0}\",", "3");
					}
					#endregion
					else if (dt.Rows[i][j].ToString().Contains(ConstDefine.END))
					{
						builder.AppendFormat("\"\"<<{0}>>", dt.Rows[i][j]);
					}//成本控制
					else if (dt.Rows[i][j].ToString().Contains("FixedCost"))
					{
						builder.AppendFormat("\"2\"<<{0}>>", ConstDefine.END);
					}
					else if (dt.Rows[i][j].ToString().Contains(ConstDefine.ZDBM))
					{

						builder.AppendFormat("\"{0}\",", string.Format(dt.Rows[i][j].ToString().Replace(ConstDefine.ZDBM, ""), (i + 1)).PadLeft(4, '0'));
					}
					else if (dt.Rows[i][j].ToString().Contains(ConstDefine.YYYYMMdd))
					{
						builder.AppendFormat("\"{0}\",", string.Format(dt.Rows[i][j].ToString().Replace(ConstDefine.YYYYMMdd, ""), doDataTime.ToString("yyyyMMdd")));
					}
					else if (dt.Rows[i][j].ToString().Contains(ConstDefine.HHmmss))
					{
						builder.AppendFormat("\"{0}\",", string.Format(dt.Rows[i][j].ToString().Replace(ConstDefine.HHmmss, ""), string.Format("{0:T}", doDataTime)));
					}
					else
					{
						builder.AppendFormat("\"{0}\",", dt.Rows[i][j]);
					}
				}
				builder.AppendFormat("\r\n");
			}

			File.WriteAllText(filePath, builder.ToString().TrimEnd(','), Encoding.Default);
			MessageBox.Show("抓取完畢！\r\n文件保存至：" + filePath);
		}
	}
}
