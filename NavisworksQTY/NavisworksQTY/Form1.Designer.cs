namespace WindowsFormsApplication7
{
    partial class Form1
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
            this.lstColumns = new System.Windows.Forms.ListBox();
            this.lstGroups = new System.Windows.Forms.ListBox();
            this.lstHeaders = new System.Windows.Forms.ListBox();
            this.button4 = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.loadQty = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.lvDataView = new System.Windows.Forms.ListView();
            this.btnRemoveHeader = new System.Windows.Forms.Button();
            this.btnAddHeader = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lstStatus = new System.Windows.Forms.ListBox();
            this.lvcLevel = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.exportCCSSheetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lstColumns
            // 
            this.lstColumns.FormattingEnabled = true;
            this.lstColumns.Location = new System.Drawing.Point(12, 47);
            this.lstColumns.Name = "lstColumns";
            this.lstColumns.Size = new System.Drawing.Size(165, 212);
            this.lstColumns.TabIndex = 2;
            // 
            // lstGroups
            // 
            this.lstGroups.FormattingEnabled = true;
            this.lstGroups.Location = new System.Drawing.Point(399, 47);
            this.lstGroups.Name = "lstGroups";
            this.lstGroups.Size = new System.Drawing.Size(120, 212);
            this.lstGroups.TabIndex = 3;
            // 
            // lstHeaders
            // 
            this.lstHeaders.FormattingEnabled = true;
            this.lstHeaders.Location = new System.Drawing.Point(228, 47);
            this.lstHeaders.Name = "lstHeaders";
            this.lstHeaders.Size = new System.Drawing.Size(165, 212);
            this.lstHeaders.TabIndex = 4;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(525, 47);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(117, 23);
            this.button4.TabIndex = 7;
            this.button4.Text = "Generate CCS Data";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1008, 24);
            this.menuStrip1.TabIndex = 10;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.loadQty,
            this.exportCCSSheetToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            this.fileToolStripMenuItem.Click += new System.EventHandler(this.fileToolStripMenuItem_Click);
            // 
            // loadQty
            // 
            this.loadQty.Name = "loadQty";
            this.loadQty.Size = new System.Drawing.Size(180, 22);
            this.loadQty.Text = "Load Take-off Sheet";
            this.loadQty.Click += new System.EventHandler(this.loadQTY_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "OFD";
            this.openFileDialog1.Filter = "(*.xlsx)|*.xlsx";
            // 
            // lvDataView
            // 
            this.lvDataView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvDataView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.lvcLevel,
            this.columnHeader1});
            this.lvDataView.GridLines = true;
            this.lvDataView.Location = new System.Drawing.Point(12, 265);
            this.lvDataView.Name = "lvDataView";
            this.lvDataView.Size = new System.Drawing.Size(984, 275);
            this.lvDataView.TabIndex = 11;
            this.lvDataView.UseCompatibleStateImageBehavior = false;
            this.lvDataView.View = System.Windows.Forms.View.Details;
            // 
            // btnRemoveHeader
            // 
            this.btnRemoveHeader.Location = new System.Drawing.Point(183, 167);
            this.btnRemoveHeader.Name = "btnRemoveHeader";
            this.btnRemoveHeader.Size = new System.Drawing.Size(39, 23);
            this.btnRemoveHeader.TabIndex = 12;
            this.btnRemoveHeader.Text = "<<";
            this.btnRemoveHeader.UseVisualStyleBackColor = true;
            this.btnRemoveHeader.Click += new System.EventHandler(this.btnRemoveHeader_Click);
            // 
            // btnAddHeader
            // 
            this.btnAddHeader.Location = new System.Drawing.Point(183, 117);
            this.btnAddHeader.Name = "btnAddHeader";
            this.btnAddHeader.Size = new System.Drawing.Size(39, 23);
            this.btnAddHeader.TabIndex = 13;
            this.btnAddHeader.Text = ">>";
            this.btnAddHeader.UseVisualStyleBackColor = true;
            this.btnAddHeader.Click += new System.EventHandler(this.btnAddHeader_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 13);
            this.label1.TabIndex = 14;
            this.label1.Text = "Available Headers:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(225, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 14;
            this.label2.Text = "Export Headers:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(396, 31);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(102, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Detected Hierarchy:";
            // 
            // lstStatus
            // 
            this.lstStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstStatus.FormattingEnabled = true;
            this.lstStatus.Location = new System.Drawing.Point(648, 47);
            this.lstStatus.Name = "lstStatus";
            this.lstStatus.Size = new System.Drawing.Size(348, 212);
            this.lstStatus.TabIndex = 0;
            // 
            // lvcLevel
            // 
            this.lvcLevel.Text = "Level";
            this.lvcLevel.Width = 50;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Bill Item";
            this.columnHeader1.Width = 250;
            // 
            // exportCCSSheetToolStripMenuItem
            // 
            this.exportCCSSheetToolStripMenuItem.Name = "exportCCSSheetToolStripMenuItem";
            this.exportCCSSheetToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.exportCCSSheetToolStripMenuItem.Text = "Export CCS Sheet";
            this.exportCCSSheetToolStripMenuItem.Click += new System.EventHandler(this.exportCCSSheetToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1008, 552);
            this.Controls.Add(this.lstStatus);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAddHeader);
            this.Controls.Add(this.btnRemoveHeader);
            this.Controls.Add(this.lvDataView);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.lstHeaders);
            this.Controls.Add(this.lstGroups);
            this.Controls.Add(this.lstColumns);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lstColumns;
        private System.Windows.Forms.ListBox lstGroups;
        private System.Windows.Forms.ListBox lstHeaders;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadQty;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ListView lvDataView;
        private System.Windows.Forms.Button btnRemoveHeader;
        private System.Windows.Forms.Button btnAddHeader;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox lstStatus;
        private System.Windows.Forms.ColumnHeader lvcLevel;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ToolStripMenuItem exportCCSSheetToolStripMenuItem;
    }
}

