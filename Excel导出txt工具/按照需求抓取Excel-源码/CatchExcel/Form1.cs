using System;
using System.Data;
using System.Windows.Forms;

namespace CatchExcel
{
	public partial class frmCatchExcel : Form
	{
		/// <summary>
		/// 保存用户选择的文件名及路径
		/// </summary>
		private string fileName = string.Empty;

		public frmCatchExcel()
		{
			InitializeComponent();
		}

		/// <summary>
		/// 获取文件名
		/// </summary>
		/// <returns></returns>
		public string GetSelectFilePathName()
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
		/// Excel導出至TXT格式文本
		/// </summary>
		/// <param name="accessSql"></param>
		/// <param name="fileNameprefix"></param>
		private void ExcelToTxt(string accessSql, string fileNameprefix)
		{
			if (string.IsNullOrEmpty(this.GetSelectFilePathName()))
			{
				return;
			}
			try
			{
				string sql = string.Empty;

				switch (fileNameprefix)
				{
					case ConstDefine.IMSDB:
					case ConstDefine.IMSEF:
						sql = string.Format(accessSql, "<>");
						break;
					default:
						sql = string.Format(accessSql, "{0}", "<>");
						break;
				}

				//if (fileNameprefix == ConstDefine.IMSDB || fileNameprefix == ConstDefine.IMSEF)
				//{
				//	sql = string.Format(accessSql, "<>");
				//}
				//else
				//{
				//	sql = string.Format(accessSql, "{0}", "<>");
				//}

				DataTable dt = CatchExcelHelper.ExcelToDataSet(fileName, sql);
				this.dgvProduct.DataSource = dt;
				string filePath = fileName.Substring(0, fileName.LastIndexOf("\\"));
				filePath += string.Format("\\{0}" + new Random().Next() + ".txt", fileNameprefix);
				//写入到txt文件中
				CatchExcelHelper.WriteTxt(filePath, dt);
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("發生錯誤，原因：{0}", ex.Message), "錯誤提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		/// <summary>
		/// 商品抓取
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnimscj_Click(object sender, EventArgs e)
		{
			ExcelToTxt(ConstDefine.AccessSqlProduct_imscj, ConstDefine.IMSCJ);
		}

		/// <summary>
		/// 库存抓取
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnImscl_Click(object sender, EventArgs e)
		{

			ExcelToTxt(ConstDefine.AccessSqlProduct_imscl, ConstDefine.IMSCL);
		}

		/// <summary>
		/// 抓取厂商
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnImsdb_Click(object sender, EventArgs e)
		{
			ExcelToTxt(ConstDefine.AccessSqlProduct_imsdb, ConstDefine.IMSDB);
		}

		/// <summary>
		/// 客戶抓取
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnImsdf_Click(object sender, EventArgs e)
		{
			ExcelToTxt(ConstDefine.AccessSqlProduct_imsef, ConstDefine.IMSEF);
		}

		/// <summary>
		/// 成本调整
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnFixedCost_Click(object sender, EventArgs e)
		{
			ExcelToTxt(ConstDefine.AccessSql_imscu, ConstDefine.IMSCU);
		}
	}
}
