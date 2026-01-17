namespace GOWordAgentAddIn
{
    partial class GOWordAgentPaneControl
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

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.txtInput = new System.Windows.Forms.TextBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnSend = new System.Windows.Forms.Button();
            this.txtChatHistory = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // txtInput
            // 
            this.txtInput.Location = new System.Drawing.Point(3, 536);
            this.txtInput.Multiline = true;
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(420, 67);
            this.txtInput.TabIndex = 2;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(3, 619);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(29, 12);
            this.lblStatus.TabIndex = 3;
            this.lblStatus.Text = "状态";
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(340, 614);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(83, 25);
            this.btnSend.TabIndex = 5;
            this.btnSend.Text = "提问";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click_1);
            // 
            // txtChatHistory
            // 
            this.txtChatHistory.Location = new System.Drawing.Point(3, 3);
            this.txtChatHistory.Name = "txtChatHistory";
            this.txtChatHistory.Size = new System.Drawing.Size(420, 527);
            this.txtChatHistory.TabIndex = 6;
            this.txtChatHistory.Text = "";
            // 
            // GOWordAgentPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.txtChatHistory);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.txtInput);
            this.Name = "GOWordAgentPaneControl";
            this.Size = new System.Drawing.Size(430, 642);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtInput;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.RichTextBox txtChatHistory;
    }
}
