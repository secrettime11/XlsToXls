namespace XlsToXls
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.sample = new System.Windows.Forms.NumericUpDown();
            this.target = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_run = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.txt_remark = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.sample)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.target)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(20, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "樣板";
            // 
            // sample
            // 
            this.sample.Location = new System.Drawing.Point(67, 24);
            this.sample.Maximum = new decimal(new int[] {
            110,
            0,
            0,
            0});
            this.sample.Name = "sample";
            this.sample.Size = new System.Drawing.Size(76, 22);
            this.sample.TabIndex = 1;
            this.sample.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.sample.Value = new decimal(new int[] {
            106,
            0,
            0,
            0});
            // 
            // target
            // 
            this.target.Location = new System.Drawing.Point(67, 56);
            this.target.Maximum = new decimal(new int[] {
            110,
            0,
            0,
            0});
            this.target.Name = "target";
            this.target.Size = new System.Drawing.Size(76, 22);
            this.target.TabIndex = 3;
            this.target.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.target.Value = new decimal(new int[] {
            105,
            0,
            0,
            0});
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(20, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "輸出";
            // 
            // btn_run
            // 
            this.btn_run.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_run.Location = new System.Drawing.Point(67, 90);
            this.btn_run.Name = "btn_run";
            this.btn_run.Size = new System.Drawing.Size(75, 27);
            this.btn_run.TabIndex = 4;
            this.btn_run.Text = "執行";
            this.btn_run.UseVisualStyleBackColor = true;
            this.btn_run.Click += new System.EventHandler(this.btn_run_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.ForeColor = System.Drawing.Color.DarkGreen;
            this.label3.Location = new System.Drawing.Point(68, 131);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 20);
            this.label3.TabIndex = 5;
            this.label3.Text = "執行狀態";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.ForeColor = System.Drawing.Color.Blue;
            this.label4.Location = new System.Drawing.Point(170, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(81, 20);
            this.label4.TabIndex = 6;
            this.label4.Text = "使用說明 :";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 10F);
            this.label5.ForeColor = System.Drawing.Color.Red;
            this.label5.Location = new System.Drawing.Point(171, 51);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(616, 18);
            this.label5.TabIndex = 7;
            this.label5.Text = "1.請先以Sample資料夾中的106.xls開始將sheet1 L欄有空白的股票填入股票代號，修改完畢後儲存";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微軟正黑體", 10F);
            this.label6.ForeColor = System.Drawing.Color.Red;
            this.label6.Location = new System.Drawing.Point(171, 75);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(672, 18);
            this.label6.TabIndex = 8;
            this.label6.Text = "2.接著設定樣板為修改完的xls名稱，輸出為接著要修改的xls名稱(請依順序 例如: 樣板=>106 則 輸出=>105)";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微軟正黑體", 10F);
            this.label7.ForeColor = System.Drawing.Color.Red;
            this.label7.Location = new System.Drawing.Point(171, 102);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(430, 18);
            this.label7.TabIndex = 9;
            this.label7.Text = "3.按下執行，執行狀態會顯示=>執行中...，執行完畢會顯示執行成功!";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微軟正黑體", 10F);
            this.label8.ForeColor = System.Drawing.Color.Red;
            this.label8.Location = new System.Drawing.Point(171, 127);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(389, 18);
            this.label8.TabIndex = 10;
            this.label8.Text = "4.確認輸出的xls檔案的修改時間有改變後，即可修改輸出的xls";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("微軟正黑體", 10F);
            this.label9.ForeColor = System.Drawing.Color.Red;
            this.label9.Location = new System.Drawing.Point(171, 153);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(439, 18);
            this.label9.TabIndex = 11;
            this.label9.Text = "5.接著重複以上動作 (樣板=>105 輸出=>104) 依序直到全部修改完畢!";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("微軟正黑體", 10F);
            this.label10.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label10.Location = new System.Drawing.Point(41, 210);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(789, 18);
            this.label10.TabIndex = 12;
            this.label10.Text = "※ 88年的xls檔sheet2有提供該年個股的代號 可以不必google 可以直接從sheet2中找到個股代號，其他都沒有就這個檔案有。";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label11.ForeColor = System.Drawing.Color.Maroon;
            this.label11.Location = new System.Drawing.Point(61, 241);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(745, 20);
            this.label11.TabIndex = 13;
            this.label11.Text = "為了避免您分次作業忘記做到哪，可留下備註在以下方格中，下次開啟程式會自動帶入上次的備註資料↓";
            // 
            // txt_remark
            // 
            this.txt_remark.Font = new System.Drawing.Font("華康儷楷書(P)", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_remark.Location = new System.Drawing.Point(43, 276);
            this.txt_remark.Multiline = true;
            this.txt_remark.Name = "txt_remark";
            this.txt_remark.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txt_remark.Size = new System.Drawing.Size(776, 118);
            this.txt_remark.TabIndex = 14;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("微軟正黑體", 10F);
            this.label12.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label12.Location = new System.Drawing.Point(171, 179);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(394, 18);
            this.label12.TabIndex = 15;
            this.label12.Text = "(提醒 : 每次執行前到執行結束前，請勿開啟樣板及輸出的xls檔)";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(852, 406);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.txt_remark);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btn_run);
            this.Controls.Add(this.target);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.sample);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "勞煩您了";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.sample)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.target)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown sample;
        private System.Windows.Forms.NumericUpDown target;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_run;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txt_remark;
        private System.Windows.Forms.Label label12;
    }
}

