namespace BST_reports
{
    partial class BSTMonitorForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.CloseButton = new System.Windows.Forms.Button();
            this.BSTFileWatcher = new System.IO.FileSystemWatcher();
            this.FileEvents = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.EventCounter = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.BSTFileWatcher)).BeginInit();
            this.SuspendLayout();
            // 
            // CloseButton
            // 
            this.CloseButton.CausesValidation = false;
            this.CloseButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CloseButton.Location = new System.Drawing.Point(129, 288);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(79, 28);
            this.CloseButton.TabIndex = 1;
            this.CloseButton.Text = "Close";
            this.CloseButton.UseVisualStyleBackColor = true;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // BSTFileWatcher
            // 
            this.BSTFileWatcher.EnableRaisingEvents = true;
            this.BSTFileWatcher.Filter = "*.htm";
            this.BSTFileWatcher.IncludeSubdirectories = true;
            this.BSTFileWatcher.NotifyFilter = System.IO.NotifyFilters.LastWrite;
            this.BSTFileWatcher.SynchronizingObject = this;
            this.BSTFileWatcher.Changed += new System.IO.FileSystemEventHandler(this.BSTFileWatcher_Changed);
            this.BSTFileWatcher.Created += new System.IO.FileSystemEventHandler(this.BSTFileWatcher_Created);
            // 
            // FileEvents
            // 
            this.FileEvents.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.FileEvents.Location = new System.Drawing.Point(31, 59);
            this.FileEvents.Multiline = true;
            this.FileEvents.Name = "FileEvents";
            this.FileEvents.ReadOnly = true;
            this.FileEvents.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.FileEvents.Size = new System.Drawing.Size(253, 199);
            this.FileEvents.TabIndex = 3;
            this.FileEvents.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 268);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Event counter:";
            // 
            // EventCounter
            // 
            this.EventCounter.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.EventCounter.Location = new System.Drawing.Point(112, 269);
            this.EventCounter.Name = "EventCounter";
            this.EventCounter.ReadOnly = true;
            this.EventCounter.Size = new System.Drawing.Size(54, 13);
            this.EventCounter.TabIndex = 6;
            this.EventCounter.TabStop = false;
            this.EventCounter.Text = "0";
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(28, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(256, 48);
            this.textBox1.TabIndex = 7;
            this.textBox1.TabStop = false;
            this.textBox1.Text = "The PrjAnalysis.htm and PrjWbs.htm files in the BST reports folder are monitored " +
    "for changes.";
            // 
            // BSTMonitorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CloseButton;
            this.ClientSize = new System.Drawing.Size(317, 327);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.EventCounter);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.FileEvents);
            this.Controls.Add(this.CloseButton);
            this.Name = "BSTMonitorForm";
            this.Text = "BSTMonitorForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.BSTMonitorForm_FormClosing);
            this.Load += new System.EventHandler(this.BSTMonitorForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.BSTFileWatcher)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button CloseButton;
        private System.IO.FileSystemWatcher BSTFileWatcher;
        private System.Windows.Forms.TextBox EventCounter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        internal System.Windows.Forms.TextBox FileEvents;
    }
}