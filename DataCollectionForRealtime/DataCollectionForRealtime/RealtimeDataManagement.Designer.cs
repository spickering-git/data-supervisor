namespace DataCollectionForRealtime
{
    partial class RealtimeDataManagement
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RealtimeDataManagement));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.statusStripOptionMonitor = new System.Windows.Forms.StatusStrip();
            this.statusOfUpdatedInstruments = new System.Windows.Forms.ToolStripStatusLabel();
            this.connectionStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusSubscribeData = new System.Windows.Forms.ToolStripStatusLabel();
            this.dataStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageExpressionsList = new System.Windows.Forms.TabPage();
            this.dataGridViewExpressionList = new System.Windows.Forms.DataGridView();
            this.tabPageSettings = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioBtnBidPriceRules = new System.Windows.Forms.RadioButton();
            this.radioBtnAskPriceRules = new System.Windows.Forms.RadioButton();
            this.radioBtnTheorPriceRules = new System.Windows.Forms.RadioButton();
            this.radioBtnMidPriceRules = new System.Windows.Forms.RadioButton();
            this.radioBtnDefaultPriceRules = new System.Windows.Forms.RadioButton();
            this.btnCallAllInstruments = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator17 = new System.Windows.Forms.ToolStripSeparator();
            this.btnCallUnsubscribed = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator18 = new System.Windows.Forms.ToolStripSeparator();
            this.btnResetAllInstruments = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator19 = new System.Windows.Forms.ToolStripSeparator();
            this.btnCQGRecon = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton3 = new System.Windows.Forms.ToolStripSeparator();
            this.mainRealtimeMenuStrip = new System.Windows.Forms.MenuStrip();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tabLog = new System.Windows.Forms.TabPage();
            this.richTextBoxLog = new System.Windows.Forms.RichTextBox();
            this.toolStripMenuItemListern = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStripOptionMonitor.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPageExpressionsList.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExpressionList)).BeginInit();
            this.tabPageSettings.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.mainRealtimeMenuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tabLog.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStripOptionMonitor
            // 
            this.statusStripOptionMonitor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusStripOptionMonitor.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusOfUpdatedInstruments,
            this.connectionStatus,
            this.statusSubscribeData,
            this.dataStatus});
            this.statusStripOptionMonitor.Location = new System.Drawing.Point(0, 391);
            this.statusStripOptionMonitor.Name = "statusStripOptionMonitor";
            this.statusStripOptionMonitor.Size = new System.Drawing.Size(795, 22);
            this.statusStripOptionMonitor.TabIndex = 6;
            this.statusStripOptionMonitor.Text = "statusStrip1";
            // 
            // statusOfUpdatedInstruments
            // 
            this.statusOfUpdatedInstruments.Name = "statusOfUpdatedInstruments";
            this.statusOfUpdatedInstruments.Size = new System.Drawing.Size(148, 17);
            this.statusOfUpdatedInstruments.Text = "statusOfUpdatedInstruments";
            // 
            // connectionStatus
            // 
            this.connectionStatus.BackColor = System.Drawing.Color.Black;
            this.connectionStatus.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.connectionStatus.ForeColor = System.Drawing.Color.White;
            this.connectionStatus.Image = ((System.Drawing.Image)(resources.GetObject("connectionStatus.Image")));
            this.connectionStatus.Name = "connectionStatus";
            this.connectionStatus.Padding = new System.Windows.Forms.Padding(0, 0, 20, 0);
            this.connectionStatus.Size = new System.Drawing.Size(114, 17);
            this.connectionStatus.Text = "CQG:WAITING";
            // 
            // statusSubscribeData
            // 
            this.statusSubscribeData.BackColor = System.Drawing.Color.White;
            this.statusSubscribeData.Name = "statusSubscribeData";
            this.statusSubscribeData.Size = new System.Drawing.Size(106, 17);
            this.statusSubscribeData.Text = "statusSubscribeData";
            // 
            // dataStatus
            // 
            this.dataStatus.BackColor = System.Drawing.Color.Black;
            this.dataStatus.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dataStatus.ForeColor = System.Drawing.Color.White;
            this.dataStatus.Name = "dataStatus";
            this.dataStatus.Size = new System.Drawing.Size(62, 17);
            this.dataStatus.Text = "dataStatus";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageExpressionsList);
            this.tabControl1.Controls.Add(this.tabPageSettings);
            this.tabControl1.Controls.Add(this.tabLog);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(559, 304);
            this.tabControl1.TabIndex = 7;
            // 
            // tabPageExpressionsList
            // 
            this.tabPageExpressionsList.Controls.Add(this.dataGridViewExpressionList);
            this.tabPageExpressionsList.Location = new System.Drawing.Point(4, 22);
            this.tabPageExpressionsList.Name = "tabPageExpressionsList";
            this.tabPageExpressionsList.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageExpressionsList.Size = new System.Drawing.Size(551, 278);
            this.tabPageExpressionsList.TabIndex = 0;
            this.tabPageExpressionsList.Text = "Expression List";
            this.tabPageExpressionsList.UseVisualStyleBackColor = true;
            // 
            // dataGridViewExpressionList
            // 
            this.dataGridViewExpressionList.AllowUserToAddRows = false;
            this.dataGridViewExpressionList.AllowUserToDeleteRows = false;
            this.dataGridViewExpressionList.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewExpressionList.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewExpressionList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewExpressionList.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewExpressionList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewExpressionList.Location = new System.Drawing.Point(3, 3);
            this.dataGridViewExpressionList.Name = "dataGridViewExpressionList";
            this.dataGridViewExpressionList.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dataGridViewExpressionList.Size = new System.Drawing.Size(545, 272);
            this.dataGridViewExpressionList.TabIndex = 6;
            // 
            // tabPageSettings
            // 
            this.tabPageSettings.Controls.Add(this.groupBox1);
            this.tabPageSettings.Location = new System.Drawing.Point(4, 22);
            this.tabPageSettings.Name = "tabPageSettings";
            this.tabPageSettings.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSettings.Size = new System.Drawing.Size(551, 278);
            this.tabPageSettings.TabIndex = 1;
            this.tabPageSettings.Text = "Settings";
            this.tabPageSettings.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioBtnBidPriceRules);
            this.groupBox1.Controls.Add(this.radioBtnAskPriceRules);
            this.groupBox1.Controls.Add(this.radioBtnTheorPriceRules);
            this.groupBox1.Controls.Add(this.radioBtnMidPriceRules);
            this.groupBox1.Controls.Add(this.radioBtnDefaultPriceRules);
            this.groupBox1.Location = new System.Drawing.Point(6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(127, 132);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Pricing Algorithm";
            // 
            // radioBtnBidPriceRules
            // 
            this.radioBtnBidPriceRules.AutoSize = true;
            this.radioBtnBidPriceRules.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioBtnBidPriceRules.Location = new System.Drawing.Point(6, 84);
            this.radioBtnBidPriceRules.Name = "radioBtnBidPriceRules";
            this.radioBtnBidPriceRules.Size = new System.Drawing.Size(39, 17);
            this.radioBtnBidPriceRules.TabIndex = 7;
            this.radioBtnBidPriceRules.Text = "Bid";
            this.radioBtnBidPriceRules.UseVisualStyleBackColor = true;
            this.radioBtnBidPriceRules.CheckedChanged += new System.EventHandler(this.radioBtnDefaultPriceRules_CheckedChanged);
            // 
            // radioBtnAskPriceRules
            // 
            this.radioBtnAskPriceRules.AutoSize = true;
            this.radioBtnAskPriceRules.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioBtnAskPriceRules.Location = new System.Drawing.Point(6, 38);
            this.radioBtnAskPriceRules.Name = "radioBtnAskPriceRules";
            this.radioBtnAskPriceRules.Size = new System.Drawing.Size(42, 17);
            this.radioBtnAskPriceRules.TabIndex = 6;
            this.radioBtnAskPriceRules.Text = "Ask";
            this.radioBtnAskPriceRules.UseVisualStyleBackColor = true;
            this.radioBtnAskPriceRules.CheckedChanged += new System.EventHandler(this.radioBtnDefaultPriceRules_CheckedChanged);
            // 
            // radioBtnTheorPriceRules
            // 
            this.radioBtnTheorPriceRules.AutoSize = true;
            this.radioBtnTheorPriceRules.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioBtnTheorPriceRules.Location = new System.Drawing.Point(6, 107);
            this.radioBtnTheorPriceRules.Name = "radioBtnTheorPriceRules";
            this.radioBtnTheorPriceRules.Size = new System.Drawing.Size(78, 17);
            this.radioBtnTheorPriceRules.TabIndex = 5;
            this.radioBtnTheorPriceRules.Text = "Theoretical";
            this.radioBtnTheorPriceRules.UseVisualStyleBackColor = true;
            this.radioBtnTheorPriceRules.CheckedChanged += new System.EventHandler(this.radioBtnDefaultPriceRules_CheckedChanged);
            // 
            // radioBtnMidPriceRules
            // 
            this.radioBtnMidPriceRules.AutoSize = true;
            this.radioBtnMidPriceRules.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioBtnMidPriceRules.Location = new System.Drawing.Point(6, 61);
            this.radioBtnMidPriceRules.Name = "radioBtnMidPriceRules";
            this.radioBtnMidPriceRules.Size = new System.Drawing.Size(80, 17);
            this.radioBtnMidPriceRules.TabIndex = 4;
            this.radioBtnMidPriceRules.Text = "Mid/Bid/Ask";
            this.radioBtnMidPriceRules.UseVisualStyleBackColor = true;
            this.radioBtnMidPriceRules.CheckedChanged += new System.EventHandler(this.radioBtnDefaultPriceRules_CheckedChanged);
            // 
            // radioBtnDefaultPriceRules
            // 
            this.radioBtnDefaultPriceRules.AutoSize = true;
            this.radioBtnDefaultPriceRules.Checked = true;
            this.radioBtnDefaultPriceRules.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioBtnDefaultPriceRules.Location = new System.Drawing.Point(6, 15);
            this.radioBtnDefaultPriceRules.Name = "radioBtnDefaultPriceRules";
            this.radioBtnDefaultPriceRules.Size = new System.Drawing.Size(60, 17);
            this.radioBtnDefaultPriceRules.TabIndex = 3;
            this.radioBtnDefaultPriceRules.TabStop = true;
            this.radioBtnDefaultPriceRules.Text = "Default";
            this.radioBtnDefaultPriceRules.UseVisualStyleBackColor = true;
            this.radioBtnDefaultPriceRules.CheckedChanged += new System.EventHandler(this.radioBtnDefaultPriceRules_CheckedChanged);
            // 
            // btnCallAllInstruments
            // 
            this.btnCallAllInstruments.BackColor = System.Drawing.Color.GreenYellow;
            this.btnCallAllInstruments.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnCallAllInstruments.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCallAllInstruments.Image = ((System.Drawing.Image)(resources.GetObject("btnCallAllInstruments.Image")));
            this.btnCallAllInstruments.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnCallAllInstruments.Name = "btnCallAllInstruments";
            this.btnCallAllInstruments.Size = new System.Drawing.Size(46, 20);
            this.btnCallAllInstruments.Text = "Call All";
            this.btnCallAllInstruments.Click += new System.EventHandler(this.btnCallAllInstruments_Click);
            // 
            // toolStripSeparator17
            // 
            this.toolStripSeparator17.Name = "toolStripSeparator17";
            this.toolStripSeparator17.Size = new System.Drawing.Size(6, 23);
            // 
            // btnCallUnsubscribed
            // 
            this.btnCallUnsubscribed.BackColor = System.Drawing.Color.Yellow;
            this.btnCallUnsubscribed.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnCallUnsubscribed.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCallUnsubscribed.Image = ((System.Drawing.Image)(resources.GetObject("btnCallUnsubscribed.Image")));
            this.btnCallUnsubscribed.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnCallUnsubscribed.Name = "btnCallUnsubscribed";
            this.btnCallUnsubscribed.Size = new System.Drawing.Size(70, 20);
            this.btnCallUnsubscribed.Text = "Call Unsub.";
            // 
            // toolStripSeparator18
            // 
            this.toolStripSeparator18.Name = "toolStripSeparator18";
            this.toolStripSeparator18.Size = new System.Drawing.Size(6, 23);
            // 
            // btnResetAllInstruments
            // 
            this.btnResetAllInstruments.BackColor = System.Drawing.Color.Red;
            this.btnResetAllInstruments.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnResetAllInstruments.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnResetAllInstruments.Image = ((System.Drawing.Image)(resources.GetObject("btnResetAllInstruments.Image")));
            this.btnResetAllInstruments.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnResetAllInstruments.Name = "btnResetAllInstruments";
            this.btnResetAllInstruments.Size = new System.Drawing.Size(37, 20);
            this.btnResetAllInstruments.Text = "Clear";
            this.btnResetAllInstruments.ToolTipText = "CLEAR ALL INSTRUMENTS";
            // 
            // toolStripSeparator19
            // 
            this.toolStripSeparator19.Name = "toolStripSeparator19";
            this.toolStripSeparator19.Size = new System.Drawing.Size(6, 23);
            // 
            // btnCQGRecon
            // 
            this.btnCQGRecon.BackColor = System.Drawing.SystemColors.Control;
            this.btnCQGRecon.ForeColor = System.Drawing.Color.Black;
            this.btnCQGRecon.Image = ((System.Drawing.Image)(resources.GetObject("btnCQGRecon.Image")));
            this.btnCQGRecon.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnCQGRecon.Name = "btnCQGRecon";
            this.btnCQGRecon.Size = new System.Drawing.Size(55, 20);
            this.btnCQGRecon.Text = "Reset";
            this.btnCQGRecon.ToolTipText = "CQG Reset";
            // 
            // toolStripButton3
            // 
            this.toolStripButton3.Name = "toolStripButton3";
            this.toolStripButton3.Size = new System.Drawing.Size(6, 23);
            // 
            // mainRealtimeMenuStrip
            // 
            this.mainRealtimeMenuStrip.BackColor = System.Drawing.SystemColors.Control;
            this.mainRealtimeMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnCallAllInstruments,
            this.toolStripSeparator17,
            this.btnCallUnsubscribed,
            this.toolStripSeparator18,
            this.btnResetAllInstruments,
            this.toolStripSeparator19,
            this.btnCQGRecon,
            this.toolStripButton3,
            this.toolStripMenuItemListern});
            this.mainRealtimeMenuStrip.Location = new System.Drawing.Point(0, 0);
            this.mainRealtimeMenuStrip.Name = "mainRealtimeMenuStrip";
            this.mainRealtimeMenuStrip.Size = new System.Drawing.Size(795, 27);
            this.mainRealtimeMenuStrip.TabIndex = 8;
            this.mainRealtimeMenuStrip.Text = "menuStrip1";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Location = new System.Drawing.Point(12, 48);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tabControl1);
            this.splitContainer1.Size = new System.Drawing.Size(771, 304);
            this.splitContainer1.SplitterDistance = 208;
            this.splitContainer1.TabIndex = 10;
            // 
            // RealtimeDataManagement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(795, 413);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.mainRealtimeMenuStrip);
            this.Controls.Add(this.statusStripOptionMonitor);
            this.Name = "RealtimeDataManagement";
            this.Text = "RealtimeDataManagement";
            this.statusStripOptionMonitor.ResumeLayout(false);
            this.statusStripOptionMonitor.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPageExpressionsList.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExpressionList)).EndInit();
            this.tabPageSettings.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.mainRealtimeMenuStrip.ResumeLayout(false);
            this.mainRealtimeMenuStrip.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tabLog.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStripOptionMonitor;
        private System.Windows.Forms.ToolStripStatusLabel statusOfUpdatedInstruments;
        private System.Windows.Forms.ToolStripStatusLabel connectionStatus;
        private System.Windows.Forms.ToolStripStatusLabel statusSubscribeData;
        private System.Windows.Forms.ToolStripStatusLabel dataStatus;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageExpressionsList;
        private System.Windows.Forms.DataGridView dataGridViewExpressionList;
        private System.Windows.Forms.TabPage tabPageSettings;
        private System.Windows.Forms.ToolStripButton btnCallAllInstruments;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator17;
        private System.Windows.Forms.ToolStripButton btnCallUnsubscribed;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator18;
        private System.Windows.Forms.ToolStripButton btnResetAllInstruments;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator19;
        private System.Windows.Forms.ToolStripButton btnCQGRecon;
        private System.Windows.Forms.ToolStripSeparator toolStripButton3;
        private System.Windows.Forms.MenuStrip mainRealtimeMenuStrip;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioBtnBidPriceRules;
        private System.Windows.Forms.RadioButton radioBtnAskPriceRules;
        private System.Windows.Forms.RadioButton radioBtnTheorPriceRules;
        private System.Windows.Forms.RadioButton radioBtnMidPriceRules;
        private System.Windows.Forms.RadioButton radioBtnDefaultPriceRules;
        private System.Windows.Forms.TabPage tabLog;
        private System.Windows.Forms.RichTextBox richTextBoxLog;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemListern;
    }
}

