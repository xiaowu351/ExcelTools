namespace NPOIOperateExcel
{
    partial class CatchExcelData
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CatchExcelData));
            this.btnA008 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dgvPart = new System.Windows.Forms.DataGridView();
            this.lbPart = new System.Windows.Forms.Label();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.lbMsg = new System.Windows.Forms.Label();
            this.dgvProduct = new System.Windows.Forms.DataGridView();
            this.lbProduct = new System.Windows.Forms.Label();
            this.btnA009 = new System.Windows.Forms.Button();
            this.btnA001BA = new System.Windows.Forms.Button();
            this.btnSaveA008 = new System.Windows.Forms.Button();
            this.btnA001KP = new System.Windows.Forms.Button();
            this.btnS003 = new System.Windows.Forms.Button();
            this.btnS003BA = new System.Windows.Forms.Button();
            this.txtSavePath = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProduct)).BeginInit();
            this.SuspendLayout();
            // 
            // btnA008
            // 
            this.btnA008.Location = new System.Drawing.Point(12, 2);
            this.btnA008.Name = "btnA008";
            this.btnA008.Size = new System.Drawing.Size(131, 28);
            this.btnA008.TabIndex = 0;
            this.btnA008.Text = "請選擇A008訂單";
            this.btnA008.UseVisualStyleBackColor = true;
            this.btnA008.Click += new System.EventHandler(this.btnA008_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dgvPart
            // 
            this.dgvPart.AllowUserToAddRows = false;
            this.dgvPart.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvPart.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPart.Location = new System.Drawing.Point(444, 103);
            this.dgvPart.Name = "dgvPart";
            this.dgvPart.RowTemplate.Height = 23;
            this.dgvPart.Size = new System.Drawing.Size(426, 283);
            this.dgvPart.TabIndex = 1;
            this.dgvPart.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
            // 
            // lbPart
            // 
            this.lbPart.AutoSize = true;
            this.lbPart.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lbPart.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.lbPart.Location = new System.Drawing.Point(440, 67);
            this.lbPart.Name = "lbPart";
            this.lbPart.Size = new System.Drawing.Size(157, 21);
            this.lbPart.TabIndex = 2;
            this.lbPart.Text = "配件數量需求表";
            // 
            // lbMsg
            // 
            this.lbMsg.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lbMsg.AutoSize = true;
            this.lbMsg.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lbMsg.ForeColor = System.Drawing.Color.Red;
            this.lbMsg.Location = new System.Drawing.Point(330, 461);
            this.lbMsg.Name = "lbMsg";
            this.lbMsg.Size = new System.Drawing.Size(0, 21);
            this.lbMsg.TabIndex = 2;
            this.lbMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // dgvProduct
            // 
            this.dgvProduct.AllowUserToAddRows = false;
            this.dgvProduct.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvProduct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvProduct.Location = new System.Drawing.Point(22, 103);
            this.dgvProduct.Name = "dgvProduct";
            this.dgvProduct.RowTemplate.Height = 23;
            this.dgvProduct.Size = new System.Drawing.Size(413, 283);
            this.dgvProduct.TabIndex = 1;
            this.dgvProduct.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dataGridView1_RowPostPaint);
            // 
            // lbProduct
            // 
            this.lbProduct.AutoSize = true;
            this.lbProduct.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lbProduct.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.lbProduct.Location = new System.Drawing.Point(18, 67);
            this.lbProduct.Name = "lbProduct";
            this.lbProduct.Size = new System.Drawing.Size(157, 21);
            this.lbProduct.TabIndex = 2;
            this.lbProduct.Text = "訂單數量需求表";
            // 
            // btnA009
            // 
            this.btnA009.Location = new System.Drawing.Point(12, 32);
            this.btnA009.Name = "btnA009";
            this.btnA009.Size = new System.Drawing.Size(131, 28);
            this.btnA009.TabIndex = 4;
            this.btnA009.Text = "請選擇A009訂單";
            this.btnA009.UseVisualStyleBackColor = true;
            this.btnA009.Click += new System.EventHandler(this.btnA009_Click);
            // 
            // btnA001BA
            // 
            this.btnA001BA.Location = new System.Drawing.Point(168, 36);
            this.btnA001BA.Name = "btnA001BA";
            this.btnA001BA.Size = new System.Drawing.Size(131, 28);
            this.btnA001BA.TabIndex = 4;
            this.btnA001BA.Text = "請選擇A001BA訂單";
            this.btnA001BA.UseVisualStyleBackColor = true;
            this.btnA001BA.Click += new System.EventHandler(this.btnA009_Click);
            // 
            // btnSaveA008
            // 
            this.btnSaveA008.BackColor = System.Drawing.SystemColors.Control;
            this.btnSaveA008.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnSaveA008.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSaveA008.Location = new System.Drawing.Point(474, 2);
            this.btnSaveA008.Name = "btnSaveA008";
            this.btnSaveA008.Size = new System.Drawing.Size(123, 62);
            this.btnSaveA008.TabIndex = 3;
            this.btnSaveA008.Text = "保存結果至Excel";
            this.btnSaveA008.UseVisualStyleBackColor = false;
            this.btnSaveA008.Click += new System.EventHandler(this.btnSaveA008_Click);
            // 
            // btnA001KP
            // 
            this.btnA001KP.Location = new System.Drawing.Point(168, 2);
            this.btnA001KP.Name = "btnA001KP";
            this.btnA001KP.Size = new System.Drawing.Size(131, 28);
            this.btnA001KP.TabIndex = 5;
            this.btnA001KP.Text = "請選擇A001KP訂單";
            this.btnA001KP.UseVisualStyleBackColor = true;
            this.btnA001KP.Click += new System.EventHandler(this.btnA001KP_Click);
            // 
            // btnS003
            // 
            this.btnS003.Location = new System.Drawing.Point(324, 2);
            this.btnS003.Name = "btnS003";
            this.btnS003.Size = new System.Drawing.Size(127, 28);
            this.btnS003.TabIndex = 6;
            this.btnS003.Text = "請選擇S003訂單";
            this.btnS003.UseVisualStyleBackColor = true;
            this.btnS003.Click += new System.EventHandler(this.btnS003_Click);
            // 
            // btnS003BA
            // 
            this.btnS003BA.Location = new System.Drawing.Point(324, 36);
            this.btnS003BA.Name = "btnS003BA";
            this.btnS003BA.Size = new System.Drawing.Size(127, 28);
            this.btnS003BA.TabIndex = 6;
            this.btnS003BA.Text = "請選擇S003BA訂單";
            this.btnS003BA.UseVisualStyleBackColor = true;
            this.btnS003BA.Click += new System.EventHandler(this.btnS003BA_Click);
            // 
            // txtSavePath
            // 
            this.txtSavePath.BackColor = System.Drawing.SystemColors.Control;
            this.txtSavePath.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtSavePath.ForeColor = System.Drawing.Color.Red;
            this.txtSavePath.Location = new System.Drawing.Point(22, 418);
            this.txtSavePath.Multiline = true;
            this.txtSavePath.Name = "txtSavePath";
            this.txtSavePath.ReadOnly = true;
            this.txtSavePath.Size = new System.Drawing.Size(848, 35);
            this.txtSavePath.TabIndex = 7;
            // 
            // CatchExcelData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(896, 463);
            this.Controls.Add(this.txtSavePath);
            this.Controls.Add(this.btnS003BA);
            this.Controls.Add(this.btnS003);
            this.Controls.Add(this.btnA001KP);
            this.Controls.Add(this.btnA001BA);
            this.Controls.Add(this.btnA009);
            this.Controls.Add(this.btnSaveA008);
            this.Controls.Add(this.lbMsg);
            this.Controls.Add(this.lbProduct);
            this.Controls.Add(this.lbPart);
            this.Controls.Add(this.dgvProduct);
            this.Controls.Add(this.dgvPart);
            this.Controls.Add(this.btnA008);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "CatchExcelData";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "康福訂單算料";
            ((System.ComponentModel.ISupportInitialize)(this.dgvPart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProduct)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnA008;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGridView dgvPart;
        private System.Windows.Forms.Label lbPart;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button btnSaveA008;
        private System.Windows.Forms.DataGridView dgvProduct;
        private System.Windows.Forms.Label lbProduct;
        public System.Windows.Forms.Label lbMsg;
        private System.Windows.Forms.Button btnA009;
        private System.Windows.Forms.Button btnA001BA;
        private System.Windows.Forms.Button btnA001KP;
        private System.Windows.Forms.Button btnS003;
        private System.Windows.Forms.Button btnS003BA;
        private System.Windows.Forms.TextBox txtSavePath;
    }
}

