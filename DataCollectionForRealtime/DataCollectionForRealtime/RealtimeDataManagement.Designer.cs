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
            this.statusStripOptionMonitor = new System.Windows.Forms.StatusStrip();
            this.statusOfUpdatedInstruments = new System.Windows.Forms.ToolStripStatusLabel();
            this.connectionStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusSubscribeData = new System.Windows.Forms.ToolStripStatusLabel();
            this.dataStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.mainRealtimeMenuStrip = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItemListern = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tabPageMongo = new System.Windows.Forms.TabPage();
            this.buttonCheck = new System.Windows.Forms.Button();
            this.buttonFill = new System.Windows.Forms.Button();
            this.richTextBoxLog = new System.Windows.Forms.RichTextBox();
            this.checkBoxAuto = new System.Windows.Forms.CheckBox();
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.statusStripOptionMonitor.SuspendLayout();
            this.mainRealtimeMenuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tabPageMongo.SuspendLayout();
            this.tabControlMain.SuspendLayout();
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
            this.connectionStatus.Size = new System.Drawing.Size(115, 17);
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
            // mainRealtimeMenuStrip
            // 
            this.mainRealtimeMenuStrip.BackColor = System.Drawing.SystemColors.Control;
            this.mainRealtimeMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemListern});
            this.mainRealtimeMenuStrip.Location = new System.Drawing.Point(0, 0);
            this.mainRealtimeMenuStrip.Name = "mainRealtimeMenuStrip";
            this.mainRealtimeMenuStrip.Size = new System.Drawing.Size(795, 24);
            this.mainRealtimeMenuStrip.TabIndex = 8;
            this.mainRealtimeMenuStrip.Text = "menuStrip1";
            // 
            // toolStripMenuItemListern
            // 
            this.toolStripMenuItemListern.Name = "toolStripMenuItemListern";
            this.toolStripMenuItemListern.Size = new System.Drawing.Size(12, 20);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Location = new System.Drawing.Point(12, 48);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tabControlMain);
            this.splitContainer1.Size = new System.Drawing.Size(771, 304);
            this.splitContainer1.SplitterDistance = 208;
            this.splitContainer1.TabIndex = 10;
            // 
            // tabPageMongo
            // 
            this.tabPageMongo.Controls.Add(this.checkBoxAuto);
            this.tabPageMongo.Controls.Add(this.richTextBoxLog);
            this.tabPageMongo.Controls.Add(this.buttonFill);
            this.tabPageMongo.Controls.Add(this.buttonCheck);
            this.tabPageMongo.Location = new System.Drawing.Point(4, 22);
            this.tabPageMongo.Name = "tabPageMongo";
            this.tabPageMongo.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageMongo.Size = new System.Drawing.Size(551, 278);
            this.tabPageMongo.TabIndex = 2;
            this.tabPageMongo.Text = "Mongo-CQG";
            this.tabPageMongo.UseVisualStyleBackColor = true;
            // 
            // buttonCheck
            // 
            this.buttonCheck.Location = new System.Drawing.Point(7, 7);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new System.Drawing.Size(166, 23);
            this.buttonCheck.TabIndex = 0;
            this.buttonCheck.Text = "Check queries in MongoDB";
            this.buttonCheck.UseVisualStyleBackColor = true;
            this.buttonCheck.Click += new System.EventHandler(this.buttonCheck_Click);
            // 
            // buttonFill
            // 
            this.buttonFill.Location = new System.Drawing.Point(179, 7);
            this.buttonFill.Name = "buttonFill";
            this.buttonFill.Size = new System.Drawing.Size(104, 23);
            this.buttonFill.TabIndex = 1;
            this.buttonFill.Text = "Fill empty data";
            this.buttonFill.UseVisualStyleBackColor = true;
            this.buttonFill.Click += new System.EventHandler(this.buttonFill_Click);
            // 
            // richTextBoxLog
            // 
            this.richTextBoxLog.Location = new System.Drawing.Point(7, 37);
            this.richTextBoxLog.Name = "richTextBoxLog";
            this.richTextBoxLog.Size = new System.Drawing.Size(538, 235);
            this.richTextBoxLog.TabIndex = 2;
            this.richTextBoxLog.Text = "";
            // 
            // checkBoxAuto
            // 
            this.checkBoxAuto.AutoSize = true;
            this.checkBoxAuto.Location = new System.Drawing.Point(290, 12);
            this.checkBoxAuto.Name = "checkBoxAuto";
            this.checkBoxAuto.Size = new System.Drawing.Size(99, 17);
            this.checkBoxAuto.TabIndex = 3;
            this.checkBoxAuto.Text = "Automatic work";
            this.checkBoxAuto.UseVisualStyleBackColor = true;
            // 
            // tabControlMain
            // 
            this.tabControlMain.Controls.Add(this.tabPageMongo);
            this.tabControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlMain.Location = new System.Drawing.Point(0, 0);
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.Size = new System.Drawing.Size(559, 304);
            this.tabControlMain.TabIndex = 7;
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
            this.Load += new System.EventHandler(this.RealtimeDataManagement_Load);
            this.statusStripOptionMonitor.ResumeLayout(false);
            this.statusStripOptionMonitor.PerformLayout();
            this.mainRealtimeMenuStrip.ResumeLayout(false);
            this.mainRealtimeMenuStrip.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tabPageMongo.ResumeLayout(false);
            this.tabPageMongo.PerformLayout();
            this.tabControlMain.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStripOptionMonitor;
        private System.Windows.Forms.ToolStripStatusLabel statusOfUpdatedInstruments;
        private System.Windows.Forms.ToolStripStatusLabel connectionStatus;
        private System.Windows.Forms.ToolStripStatusLabel statusSubscribeData;
        private System.Windows.Forms.ToolStripStatusLabel dataStatus;
        private System.Windows.Forms.MenuStrip mainRealtimeMenuStrip;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemListern;
        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TabPage tabPageMongo;
        private System.Windows.Forms.CheckBox checkBoxAuto;
        private System.Windows.Forms.RichTextBox richTextBoxLog;
        private System.Windows.Forms.Button buttonFill;
        private System.Windows.Forms.Button buttonCheck;
    }
}

