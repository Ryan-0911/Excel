namespace Excel
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
            this.btn讀取單一儲存格 = new System.Windows.Forms.Button();
            this.btn寫入單一儲存格 = new System.Windows.Forms.Button();
            this.btn讀取多重儲存格 = new System.Windows.Forms.Button();
            this.btn寫入多重儲存格 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btn讀取單一儲存格
            // 
            this.btn讀取單一儲存格.Location = new System.Drawing.Point(99, 104);
            this.btn讀取單一儲存格.Margin = new System.Windows.Forms.Padding(2);
            this.btn讀取單一儲存格.Name = "btn讀取單一儲存格";
            this.btn讀取單一儲存格.Size = new System.Drawing.Size(128, 52);
            this.btn讀取單一儲存格.TabIndex = 0;
            this.btn讀取單一儲存格.Text = "讀取單一儲存格";
            this.btn讀取單一儲存格.UseVisualStyleBackColor = true;
            this.btn讀取單一儲存格.Click += new System.EventHandler(this.btn讀取單一儲存格_Click);
            // 
            // btn寫入單一儲存格
            // 
            this.btn寫入單一儲存格.Location = new System.Drawing.Point(328, 104);
            this.btn寫入單一儲存格.Margin = new System.Windows.Forms.Padding(2);
            this.btn寫入單一儲存格.Name = "btn寫入單一儲存格";
            this.btn寫入單一儲存格.Size = new System.Drawing.Size(128, 52);
            this.btn寫入單一儲存格.TabIndex = 1;
            this.btn寫入單一儲存格.Text = "寫入單一儲存格";
            this.btn寫入單一儲存格.UseVisualStyleBackColor = true;
            this.btn寫入單一儲存格.Click += new System.EventHandler(this.btn寫入單一儲存格_Click);
            // 
            // btn讀取多重儲存格
            // 
            this.btn讀取多重儲存格.Location = new System.Drawing.Point(99, 209);
            this.btn讀取多重儲存格.Margin = new System.Windows.Forms.Padding(2);
            this.btn讀取多重儲存格.Name = "btn讀取多重儲存格";
            this.btn讀取多重儲存格.Size = new System.Drawing.Size(128, 52);
            this.btn讀取多重儲存格.TabIndex = 4;
            this.btn讀取多重儲存格.Text = "讀取多重儲存格";
            this.btn讀取多重儲存格.UseVisualStyleBackColor = true;
            this.btn讀取多重儲存格.Click += new System.EventHandler(this.btn讀取多重儲存格_Click);
            // 
            // btn寫入多重儲存格
            // 
            this.btn寫入多重儲存格.Location = new System.Drawing.Point(328, 209);
            this.btn寫入多重儲存格.Margin = new System.Windows.Forms.Padding(2);
            this.btn寫入多重儲存格.Name = "btn寫入多重儲存格";
            this.btn寫入多重儲存格.Size = new System.Drawing.Size(128, 52);
            this.btn寫入多重儲存格.TabIndex = 5;
            this.btn寫入多重儲存格.Text = "寫入多重儲存格";
            this.btn寫入多重儲存格.UseVisualStyleBackColor = true;
            this.btn寫入多重儲存格.Click += new System.EventHandler(this.btn寫入多重儲存格_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Excel.Properties.Resources.excel_ms_5bfc379146e0fb00511cdefe;
            this.pictureBox1.Location = new System.Drawing.Point(642, 46);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(154, 110);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(871, 500);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btn寫入多重儲存格);
            this.Controls.Add(this.btn讀取多重儲存格);
            this.Controls.Add(this.btn寫入單一儲存格);
            this.Controls.Add(this.btn讀取單一儲存格);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn讀取單一儲存格;
        private System.Windows.Forms.Button btn寫入單一儲存格;
        private System.Windows.Forms.Button btn讀取多重儲存格;
        private System.Windows.Forms.Button btn寫入多重儲存格;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

