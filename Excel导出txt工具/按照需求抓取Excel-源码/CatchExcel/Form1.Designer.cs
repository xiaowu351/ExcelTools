namespace CatchExcel
{
    partial class frmCatchExcel
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCatchExcel));
            this.btnimscj = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dgvProduct = new System.Windows.Forms.DataGridView();
            this.btnImsco = new System.Windows.Forms.Button();
            this.btnImsdb = new System.Windows.Forms.Button();
            this.btnImsef = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProduct)).BeginInit();
            this.SuspendLayout();
            // 
            // btnimscj
            // 
            this.btnimscj.Location = new System.Drawing.Point(15, 26);
            this.btnimscj.Name = "btnimscj";
            this.btnimscj.Size = new System.Drawing.Size(85, 60);
            this.btnimscj.TabIndex = 0;
            this.btnimscj.Text = "商品基本資料";
            this.btnimscj.UseVisualStyleBackColor = true;
            this.btnimscj.Click += new System.EventHandler(this.btnimscj_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dgvProduct
            // 
            this.dgvProduct.AllowUserToAddRows = false;
            this.dgvProduct.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvProduct.Location = new System.Drawing.Point(31, 114);
            this.dgvProduct.Name = "dgvProduct";
            this.dgvProduct.RowTemplate.Height = 23;
            this.dgvProduct.Size = new System.Drawing.Size(330, 33);
            this.dgvProduct.TabIndex = 2;
            this.dgvProduct.Visible = false;
            // 
            // btnImsco
            // 
            this.btnImsco.Location = new System.Drawing.Point(112, 26);
            this.btnImsco.Name = "btnImsco";
            this.btnImsco.Size = new System.Drawing.Size(85, 60);
            this.btnImsco.TabIndex = 3;
            this.btnImsco.Text = "庫存資料";
            this.btnImsco.UseVisualStyleBackColor = true;
            this.btnImsco.Click += new System.EventHandler(this.btnImscl_Click);
            // 
            // btnImsdb
            // 
            this.btnImsdb.Location = new System.Drawing.Point(207, 26);
            this.btnImsdb.Name = "btnImsdb";
            this.btnImsdb.Size = new System.Drawing.Size(85, 60);
            this.btnImsdb.TabIndex = 4;
            this.btnImsdb.Text = "廠商資料";
            this.btnImsdb.UseVisualStyleBackColor = true;
            this.btnImsdb.Click += new System.EventHandler(this.btnImsdb_Click);
            // 
            // btnImsef
            // 
            this.btnImsef.Location = new System.Drawing.Point(303, 26);
            this.btnImsef.Name = "btnImsef";
            this.btnImsef.Size = new System.Drawing.Size(85, 60);
            this.btnImsef.TabIndex = 4;
            this.btnImsef.Text = "客戶資料";
            this.btnImsef.UseVisualStyleBackColor = true;
            this.btnImsef.Click += new System.EventHandler(this.btnImsdf_Click);
            // 
            // frmCatchExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::CatchExcel.Properties.Resources._11;
            this.ClientSize = new System.Drawing.Size(400, 155);
            this.Controls.Add(this.btnImsef);
            this.Controls.Add(this.btnImsdb);
            this.Controls.Add(this.btnImsco);
            this.Controls.Add(this.btnimscj);
            this.Controls.Add(this.dgvProduct);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmCatchExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "鼎新电脑导入工具";
            ((System.ComponentModel.ISupportInitialize)(this.dgvProduct)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnimscj;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGridView dgvProduct;
        private System.Windows.Forms.Button btnImsco;
        private System.Windows.Forms.Button btnImsdb;
        private System.Windows.Forms.Button btnImsef;
    }
}

