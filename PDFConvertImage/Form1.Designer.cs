namespace PDFConvertImage
{
    partial class PDFtoImage
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button_LoadFile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button_LoadFile
            // 
            this.button_LoadFile.Location = new System.Drawing.Point(12, 11);
            this.button_LoadFile.Name = "button_LoadFile";
            this.button_LoadFile.Size = new System.Drawing.Size(75, 23);
            this.button_LoadFile.TabIndex = 0;
            this.button_LoadFile.Text = "載入檔案";
            this.button_LoadFile.UseVisualStyleBackColor = true;
            this.button_LoadFile.Click += new System.EventHandler(this.button_LoadFile_Click);
            // 
            // PDFtoImage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(116, 43);
            this.Controls.Add(this.button_LoadFile);
            this.Name = "PDFtoImage";
            this.Text = "PDF Convert to Image";
            this.Load += new System.EventHandler(this.PDFtoImage_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button_LoadFile;
    }
}

