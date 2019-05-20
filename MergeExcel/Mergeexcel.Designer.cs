namespace MergeExcel
{
    partial class Mergeexcel
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
            this.btnOpen = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnOpreat = new System.Windows.Forms.Button();
            this.openPath = new System.Windows.Forms.TextBox();
            this.savePath = new System.Windows.Forms.TextBox();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnOpen.Location = new System.Drawing.Point(13, 13);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(75, 25);
            this.btnOpen.TabIndex = 0;
            this.btnOpen.Text = "打开路径";
            this.btnOpen.UseMnemonic = false;
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(13, 81);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 22);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "保存路径";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnOpreat
            // 
            this.btnOpreat.Location = new System.Drawing.Point(13, 146);
            this.btnOpreat.Name = "btnOpreat";
            this.btnOpreat.Size = new System.Drawing.Size(75, 38);
            this.btnOpreat.TabIndex = 2;
            this.btnOpreat.Text = "开始操作";
            this.btnOpreat.UseVisualStyleBackColor = true;
            this.btnOpreat.Click += new System.EventHandler(this.btnOpreat_Click);
            // 
            // openPath
            // 
            this.openPath.Location = new System.Drawing.Point(94, 14);
            this.openPath.Multiline = true;
            this.openPath.Name = "openPath";
            this.openPath.Size = new System.Drawing.Size(257, 24);
            this.openPath.TabIndex = 3;
            // 
            // savePath
            // 
            this.savePath.Location = new System.Drawing.Point(94, 81);
            this.savePath.Multiline = true;
            this.savePath.Name = "savePath";
            this.savePath.Size = new System.Drawing.Size(257, 22);
            this.savePath.TabIndex = 4;
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // Mergeexcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(363, 208);
            this.Controls.Add(this.savePath);
            this.Controls.Add(this.openPath);
            this.Controls.Add(this.btnOpreat);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnOpen);
            this.Name = "Mergeexcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "合并表格";
            this.Load += new System.EventHandler(this.mergeexcel_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnOpreat;
        private System.Windows.Forms.TextBox openPath;
        private System.Windows.Forms.TextBox savePath;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnOpen;
    }
}

