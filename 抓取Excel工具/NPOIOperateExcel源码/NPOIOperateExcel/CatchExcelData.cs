using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Eval;
using System.Data.OleDb;
using NPOI.HSSF.Record;
using CommonFunction;

namespace NPOIOperateExcel
{
    public partial class CatchExcelData : Form
    {
        public CatchExcelData()
        {
            
            InitializeComponent(); 
        }
        /// <summary>
        /// 保存读取出来的A008订单需求数据
        /// </summary> 
        protected DataTable dTableA008Product;
        /// <summary>
        /// 保存读取出来的A008配件需求数据
        /// </summary> 
        protected DataTable dTableA008Part;
        /// <summary>
        /// 保存读取出来的A009组合零件需求数据
        /// </summary>
        protected DataTable dTableA009ComponentParts;
        /// <summary>
        /// 保存读取出来的A009配件需求数据
        /// </summary>
        protected DataTable dTableA009Parts;
        /// <summary>
        /// 保存读取出来的A001KP配件需求数据部分1，包括螺丝包，KP-算料，及大把手算料的前小部分
        /// </summary>
        protected DataTable dTableA001KPParts1;
        /// <summary>
        /// 保存读取出来的A001KP配件需求数据部分2，包括大把手算料的后大部分，及白紅銅-合頁算料全部
        /// </summary>
        protected DataTable dTableA001KPParts2;
        /// <summary>
        /// 保存读取出来的A001KP订单需求数据
        /// </summary>
        protected DataTable dTableA001KPProduct;
        /// <summary>
        /// 保存读取出来的S003BA订单需求数据
        /// </summary>
        protected DataTable dTableS003BAProduct;
        /// <summary>
        /// 保存读取出来的S003BA配件需求数据
        /// </summary>
        protected DataTable dTableS003BAPart;
        /// <summary>
        /// 保存读取出来的S003订单需求数据
        /// </summary>
        protected DataTable dTableS003Product;
        /// <summary>
        /// 保存读取出来的S003BA配件需求数据
        /// </summary>
        protected DataTable dTableS003Part;
        /// <summary>
        /// 保存用户选择的文件名及路径
        /// </summary>
        string fileName = string.Empty;
        /// <summary>
        /// 选择A008訂單的单击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnA008_Click(object sender, EventArgs e)
        {
            //清空顯示數據
            this.dgvProduct.Visible = true;
            this.dgvProduct.DataSource = null;
            this.dgvPart.DataSource = null;
            // 1.获取打开的Excel文件 
            this.openFileDialog1.Filter = "Excel文件|*.xls;*.xlsx;";
            if (this.openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }
            fileName = this.openFileDialog1.FileName;
            this.txtSavePath.Text = "正在努力為您抓取数据中...请稍后！";
            this.txtSavePath.Refresh();
            //this.lbMsg.Text = "正在努力為您抓取数据中...请稍后！";
            //this.lbMsg.Refresh();  // 这句是立刻显示赋值的文本内容，而不是等程序执行完成后再显示
            //System.Threading.Thread.Sleep(1000);
            #region MyRegion 
            try
            {
                // 从配置文件中读取订单部分 

                string strExcel = AppConfigHelper.Str008SqlProduct;
                dTableA008Product =ExcelHelper.ExcelToDataSet(fileName, strExcel);
                //StringBuilder sb = CreateAccessSql();
                // 从配置文件中读取配料部分
                strExcel = AppConfigHelper.Str008SqlPart;
                dTableA008Part =ExcelHelper.ExcelToDataSet(fileName, strExcel);
                // 构造含旧编号的
                // sb = CreateAccessSqlOld();
                // dTable2 = ExcelToDataSet(fileName, sb.ToString());
                if (dTableA008Product != null && dTableA008Part != null)
                {
                    // 配料需求表
                    this.dgvPart.DataSource = dTableA008Part;
                    dgvPart.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dgvPart.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dgvPart.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dgvPart.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                    // 订单数量需求表
                    this.dgvProduct.DataSource = dTableA008Product;
                    dgvProduct.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dgvProduct.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dgvProduct.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                    dgvProduct.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                    //dataGridView3.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
                    this.txtSavePath.Text = "数据抓取完毕!";
                    // MessageBox.Show("数据抓取完毕！"); 
                }
                else
                {
                    MessageBox.Show("数据抓取失败!");
                    this.txtSavePath.Text = "数据抓取失败!";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            } 
            #endregion
        }  
         
        /// <summary>
        /// 绘制生成dataGridview表格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
         private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
         {
             var dgv = sender as DataGridView;
             if (dgv != null)
             {
                 Rectangle rect = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y, dgv.RowHeadersWidth - 4, e.RowBounds.Height);
                 TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(), dgv.RowHeadersDefaultCellStyle.Font, rect, dgv.RowHeadersDefaultCellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
             }
         }

        /// <summary>
        /// 保存A008文件的单击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
         private void btnSaveA008_Click(object sender, EventArgs e)
         {
             string filePath = string.Empty;
             this.saveFileDialog1.Filter = "Excel文件|*.xls;*.xlsx;";
             if (this.saveFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
             {
                 return;
             }
             filePath = this.saveFileDialog1.FileName;
             if (string.IsNullOrEmpty(filePath))
             {
                 MessageBox.Show("请输入文件名");
                 return; 
             }
             this.txtSavePath.Text = "正在努力为您写入本地磁盘中...请稍后！";
             this.txtSavePath.Refresh();
             //this.lbMsg.Text = "正在努力为您写入本地磁盘中...请稍后！";
             //this.lbMsg.Refresh();
             string name = Path.GetFileNameWithoutExtension(filePath);
             string directory = Path.GetDirectoryName(filePath);
             try
             {
                 int count = 0;
                 #region 保存A008的部分
                 if (dTableA008Product != null && dTableA008Part != null)
                 {
                     count++; // 记录保存的个数
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableA008Product, "A008訂單數量需求表", directory + @"\A008訂單數量需求表.xls");
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableA008Part, "A008配件數量需求表", directory + @"\A008配件數量需求表.xls");
                     this.txtSavePath.Text = "文件成功保存至：" + directory;
                     //this.lbMsg.Text = "保存成功！";
                     //return;
                 }

                 #endregion

                 #region 保存A009和A001BA的部分
                 if (dTableA009ComponentParts != null && dTableA009Parts != null)
                 {
                     count++;
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableA009ComponentParts, "组合零件數量需求表", directory + @"\A009_A001BA组合零件數量需求表.xls");
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableA009Parts, "配件數量需求表", directory + @"\A009_A001BA配件數量需求表.xls");
                     this.txtSavePath.Text = "文件成功保存至：" + directory;
                     //this.lbMsg.Text = "保存成功！";
                     //return;
                 }
                 #endregion

                 #region 保存A001KP的抓取数据
                 if (dTableA001KPParts1 != null && dTableA001KPParts2 != null && dTableA001KPProduct != null)
                 {
                     count++;
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableA001KPProduct, "訂單數量需求表", directory + @"\A001KP訂單數量需求表.xls");
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableA001KPParts1, "配件數量需求表1", directory + @"\A001KP配件數量需求表1.xls");
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableA001KPParts2, "配件數量需求表2", directory + @"\A001KP配件數量需求表2.xls");
                     this.txtSavePath.Text = "文件成功保存至：" + directory;
                     //this.lbMsg.Text = "保存成功！";
                     //return;
                 }
                 #endregion

                 #region 保存S003BA抓取数据
                 if (dTableS003BAProduct != null && dTableS003BAPart != null)
                 {
                     count++;
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableS003BAProduct, "S003BA訂單數量需求表", directory + @"\S003BA訂單數量需求表.xls");
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableS003BAPart, "S003BA配件數量需求表", directory + @"\S003BA配件數量需求表.xls");

                     this.txtSavePath.Text = "文件成功保存至：" + directory;
                     //this.lbMsg.Text = "保存成功！";
                     //return;
                 }
                 #endregion

                 #region 保存S003抓取数据dTableS003Product != null &&
                 if (dTableS003Part != null)
                 {
                     count++;
                     //CommonFunction.ExcelHelper.ExportDTtoExcel(dTableS003Product, "S003訂單數量需求表", directory + @"\S003訂單數量需求表.xls");
                     CommonFunction.ExcelHelper.ExportDTtoExcel(dTableS003Part, "S003配件數量需求表", directory + @"\S003配件數量需求表.xls");

                     this.txtSavePath.Text = "文件成功保存至：" + directory;
                     //this.lbMsg.Text = "保存成功！";
                     //return;
                 }
                 #endregion

                 if (count == 0)
                 {
                     MessageBox.Show("文件写入失败，没有任何内容可写！");
                     this.txtSavePath.Text = "文件写入失败，没有任何内容可写！";
                     return;
                 } 
             }
             catch (Exception ex)
             {

                 MessageBox.Show(ex.Message);
             } 
         }
         
        /// <summary>
        /// A009或者A001BA的单击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
         private void btnA009_Click(object sender, EventArgs e)
         {
             //清空顯示數據
             this.dgvProduct.Visible = true;
             this.dgvProduct.DataSource = null;
             this.dgvPart.DataSource = null;
             if (this.openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
             {
                 return;
             }
             this.fileName = openFileDialog1.FileName;
             if (string.IsNullOrEmpty(fileName))
             {
                 MessageBox.Show("请先选择文件");
                 return;
             }
             this.txtSavePath.Text = "正在努力為您抓取数据中...请稍后！";
             this.txtSavePath.Refresh();
             //this.lbMsg.Text = "正在努力為您抓取数据中...请稍后！";
             //this.lbMsg.Refresh();

             #region MyRegion
             try
             {
                 // 从配置文件中读取订单部分 
                 string sqlExcel = AppConfigHelper.Str009SqlProduct;
                 this.lbProduct.Text = "組合零件數量需求表";
                 this.lbProduct.Refresh();// 立刻显示
                 this.dTableA009ComponentParts = ExcelHelper.ExcelToDataSet(this.fileName, sqlExcel);
                 // 读取配件部分的字符串
                 sqlExcel = AppConfigHelper.Str009SqlPart;
                 this.dTableA009Parts = ExcelHelper.ExcelToDataSet(fileName, sqlExcel);
                 if (dTableA009ComponentParts != null && dTableA009Parts!=null )
                 { 
                     // 配件需求表
                     this.dgvPart.DataSource = dTableA009Parts;
                     dgvPart.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                     // 组合零件数量需求表
                     this.dgvProduct.DataSource = this.dTableA009ComponentParts;
                     dgvProduct.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable; 
                     //this.lbMsg.Text = "数据抓取完毕!";
                     this.txtSavePath.Text = "数据抓取完毕！";
                      
                 }
                 else
                 {
                     this.txtSavePath.Text = "数据抓取失败！";
                     //this.lbMsg.Text = "数据抓取失败";
                     //this.lbMsg.Refresh();// 立刻显示
                 }
             }
             catch (Exception ex)
             { 
                 MessageBox.Show(ex.Message);
             }
             #endregion 
         }
         /// <summary>
         /// A001KP訂單的单击处理程序
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
         private void btnA001KP_Click(object sender, EventArgs e)
         {
             //清空顯示數據
             this.dgvProduct.Visible = true;
             this.dgvProduct.DataSource = null;
             this.dgvPart.DataSource = null;
             // 1.获取打开的Excel文件 
             this.openFileDialog1.Filter = "Excel文件|*.xls;*.xlsx;";
             if (this.openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
             {
                 return;
             }
             fileName = this.openFileDialog1.FileName;
             this.txtSavePath.Text = "正在努力為您抓取数据中...请稍后！";
             this.txtSavePath.Refresh();
             //this.lbMsg.Text = "正在努力為您抓取数据中...请稍后！";
             //this.lbMsg.Refresh(); 
             try
             {
                 // 从配置文件中读取配料部分 AppConfigHelper.StrA001KPSqlPart1+AppConfigHelper.StrA001KPSqlPart2 
                 //AppConfigHelper.StrA001KPSqlPart3 + AppConfigHelper.StrA001KPSqlPart4+ AppConfigHelper.StrA001KPSqlPart2
                 StringBuilder strExcelBuilder = new StringBuilder();
                 strExcelBuilder.Append(AppConfigHelper.StrA001KPSqlPart1 + AppConfigHelper.StrA001KPSqlPart2);
                 dTableA001KPParts1 = ExcelHelper.ExcelToDataSet(fileName, strExcelBuilder.ToString());
                 strExcelBuilder.Remove(0, strExcelBuilder.Length).Append(AppConfigHelper.StrA001KPSqlPart3 + AppConfigHelper.StrA001KPSqlPart4);
                 dTableA001KPParts2 = ExcelHelper.ExcelToDataSet(fileName, strExcelBuilder.ToString());
                 strExcelBuilder.Remove(0, strExcelBuilder.Length).Append(AppConfigHelper.StrA001KPSqlProduct);
                 dTableA001KPProduct = ExcelHelper.ExcelToDataSet(fileName, strExcelBuilder.ToString());
                 // && dTableA001KPProduct != null && dTableA001KPParts2 != null
                 if (dTableA001KPParts1 != null && dTableA001KPProduct != null && dTableA001KPParts2 != null)
                 {


                     // 配料需求表  
                     this.dgvPart.DataSource = dTableA001KPParts1;
                     dgvPart.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                     // 订单数量需求表
                     this.dgvProduct.DataSource = dTableA001KPProduct;
                     dgvProduct.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable; 
                     this.txtSavePath.Text = "数据抓取完毕!";
                     MessageBox.Show("数据抓取完毕！"); 
                 }
                 else
                 {
                     MessageBox.Show("数据抓取失败!");
                     this.txtSavePath.Text = "数据抓取失败!";
                 }
             }
             catch (Exception ex)
             {

                 MessageBox.Show(ex.Message);
             } 
         }
         /// <summary>
         /// S003訂單的抓取 处理程序
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
         private void btnS003_Click(object sender, EventArgs e)
         {
             //清空顯示數據
             this.dgvProduct.Visible = false;
             this.dgvPart.DataSource = null; 
             // 1.获取打开的Excel文件 
             this.openFileDialog1.Filter = "Excel文件|*.xls;*.xlsx;";
             if (this.openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
             {
                 return;
             }
             fileName = this.openFileDialog1.FileName;
             this.txtSavePath.Text = "正在努力為您抓取数据中...请稍后！";
             this.txtSavePath.Refresh();
             try
             {
                 // 从配置文件中读取配料部分 AppConfigHelper.StrA001KPSqlPart1+AppConfigHelper.StrA001KPSqlPart2 
                 //AppConfigHelper.StrA001KPSqlPart3 + AppConfigHelper.StrA001KPSqlPart4
                 StringBuilder strExcelBuilder = new StringBuilder();
                 //strExcelBuilder.Append(AppConfigHelper.StrS003BASqlProduct);
                 //dTableS003BAProduct = ExcelHelper.ExcelToDataSet(fileName, strExcelBuilder.ToString());
                 strExcelBuilder.Remove(0, strExcelBuilder.Length).Append(AppConfigHelper.StrS003SqlPart);
                 dTableS003Part = ExcelHelper.ExcelToDataSet(fileName, strExcelBuilder.ToString());


                 if (dTableS003Part != null )
                 {


                     // 配料需求表  && dTableS003Part != null
                     this.dgvPart.DataSource = dTableS003Part;
                     dgvPart.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                     // 订单数量需求表
                     //this.dgvProduct.DataSource = dTableS003BAProduct;
                     //dgvProduct.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                     //dgvProduct.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                     //dgvProduct.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                     //dgvProduct.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                     this.txtSavePath.Text = "数据抓取完毕!";
                     // MessageBox.Show("数据抓取完毕！"); 
                 }
                 else
                 {
                     MessageBox.Show("数据抓取失败!");
                     this.txtSavePath.Text = "数据抓取失败!";
                 }
             }
             catch (Exception ex)
             {

                 MessageBox.Show(ex.Message);
             } 
         }
         /// <summary>
         /// S003BA抓取 处理程序
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
         private void btnS003BA_Click(object sender, EventArgs e)
         {
             //清空顯示數據
             this.dgvProduct.Visible = true;
             this.dgvProduct.DataSource = null;
             this.dgvPart.DataSource = null;
             // 1.获取打开的Excel文件 
             this.openFileDialog1.Filter = "Excel文件|*.xls;*.xlsx;";
             if (this.openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
             {
                 return;
             }
             fileName = this.openFileDialog1.FileName;
             this.txtSavePath.Text = "正在努力為您抓取数据中...请稍后！";
             //this.lbMsg.Text = "正在努力為您抓取数据中...请稍后！";
             this.txtSavePath.Refresh();
             try
             {
                 // 从配置文件中读取配料部分 AppConfigHelper.StrA001KPSqlPart1+AppConfigHelper.StrA001KPSqlPart2 
                 //AppConfigHelper.StrA001KPSqlPart3 + AppConfigHelper.StrA001KPSqlPart4
                 StringBuilder strExcelBuilder = new StringBuilder();
                 strExcelBuilder.Append(AppConfigHelper.StrS003BASqlProduct);
                 dTableS003BAProduct = ExcelHelper.ExcelToDataSet(fileName, strExcelBuilder.ToString());
                 strExcelBuilder.Remove(0, strExcelBuilder.Length).Append(AppConfigHelper.StrS003BASqlPart);
                 dTableS003BAPart = ExcelHelper.ExcelToDataSet(fileName, strExcelBuilder.ToString());


                 if (dTableS003BAProduct != null && dTableS003BAPart != null  )
                 {


                     // 配料需求表  
                     this.dgvPart.DataSource = dTableS003BAPart;
                     dgvPart.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvPart.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                     // 订单数量需求表
                     this.dgvProduct.DataSource = dTableS003BAProduct;
                     dgvProduct.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
                     dgvProduct.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
                     this.txtSavePath.Text = "数据抓取完毕!";
                     // MessageBox.Show("数据抓取完毕！"); 
                 }
                 else
                 {
                     MessageBox.Show("数据抓取失败!");
                     this.txtSavePath.Text = "数据抓取失败!";
                 }
             }
             catch (Exception ex)
             {

                 MessageBox.Show(ex.Message);
             } 
         }

          
    }
}



/*
 *螺絲K,規格說明L,用量M,總用量N,安全庫存量O
 *塑膠R,	規格說明S	,用量T,	總用量U	,安全庫存量V        --   NOT   (總用量BD  =   '0')  作为条件
 *膠膜Y,	規格說明Z,	用量AA,	總用量AB,	安全庫存量AC
 *密著卡AF,	規格說明AG,	用量AH,	總用量AI,	安全庫存量AJ
 *真空格AM,	規格說明AN,	用量AO,	總用量AP,	安全庫存量AQ
 *真空格AT,	規格說明AU,	用量AV,	總用量AW,	安全庫存量AX
 *氣泡袋BA,	規格說明BB,	用量BC,	總用量BD,	安全庫存量BE    -- 总数量>0  作为条件
 *舒服袋BH,	規格說明BI,	用量BJ,	總用量BK,	安全庫存量BL    -- 总数量>0  作为条件
 *塑膠袋BO,	規格說明BP,	用量BQ,	總用量BR,	安全庫存量BS
 *插卡BV,	規格說明BW,	用量BX,	總用量BY,	安全庫存量BZ
 *電腦條碼紙CC,	規格說明CD,	用量CE,	總用量CF,	安全庫存量CG
 *
 * -----------------含有旧编号-----------
 *天地板CJ,	天地板舊編號CK,	規格說明CL,	用量CM,總用量CN	,安全庫存量CO
 *內盒CR,	內盒舊編號CS,	規格說明CT,	用量CU,	總用量CV,	安全庫存量CW  -- 总数量>0  作为条件
 *外箱CZ,	外箱舊編號DA,	規格說明DB,	用量DC,	總用量DD,	安全庫存量DE  -- 总数量>0  作为条件
 
     sb.AppendFormat(@" select top 1 '舒服袋' as 零件名称, '規格說明' as 规格说明,'' as 用量,'' as 总用量,'' as 安全库存量 from [算料$];UNION ALL \r\n");
            sb.AppendFormat("select 舒服袋BH ,規格說明BI ,用量BJ ,總用量BK ,安全庫存量BL  from [算料$]  where  總用量BK>0; \r\nUNION ALL \r\n"); 
*/