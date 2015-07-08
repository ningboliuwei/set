namespace ExceltkGUI
{
	partial class ExceltkGUI
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
			this.components = new System.ComponentModel.Container();
			this.menuStripMain = new System.Windows.Forms.MenuStrip();
			this.文件FToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.OpenFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.ExitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.ofdExcel = new System.Windows.Forms.OpenFileDialog();
			this.tmrProgressBar = new System.Windows.Forms.Timer(this.components);
			this.bgwImageRecognition = new System.ComponentModel.BackgroundWorker();
			this.rtxCode = new System.Windows.Forms.RichTextBox();
			this.lvwFiles = new System.Windows.Forms.ListView();
			this.statusStrip1 = new System.Windows.Forms.StatusStrip();
			this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
			this.menuStripMain.SuspendLayout();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.statusStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// menuStripMain
			// 
			this.menuStripMain.ImageScalingSize = new System.Drawing.Size(24, 24);
			this.menuStripMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文件FToolStripMenuItem});
			this.menuStripMain.Location = new System.Drawing.Point(0, 0);
			this.menuStripMain.Name = "menuStripMain";
			this.menuStripMain.Size = new System.Drawing.Size(978, 32);
			this.menuStripMain.TabIndex = 0;
			this.menuStripMain.Text = "menuStrip1";
			// 
			// 文件FToolStripMenuItem
			// 
			this.文件FToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.OpenFileToolStripMenuItem,
            this.toolStripSeparator1,
            this.ExitToolStripMenuItem});
			this.文件FToolStripMenuItem.Name = "文件FToolStripMenuItem";
			this.文件FToolStripMenuItem.Size = new System.Drawing.Size(80, 28);
			this.文件FToolStripMenuItem.Text = "文件(&F)";
			// 
			// OpenFileToolStripMenuItem
			// 
			this.OpenFileToolStripMenuItem.Name = "OpenFileToolStripMenuItem";
			this.OpenFileToolStripMenuItem.Size = new System.Drawing.Size(156, 30);
			this.OpenFileToolStripMenuItem.Text = "打开(&O)";
			this.OpenFileToolStripMenuItem.Click += new System.EventHandler(this.OpenFileToolStripMenuItem_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(153, 6);
			// 
			// ExitToolStripMenuItem
			// 
			this.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem";
			this.ExitToolStripMenuItem.Size = new System.Drawing.Size(156, 30);
			this.ExitToolStripMenuItem.Text = "退出(&X)";
			this.ExitToolStripMenuItem.Click += new System.EventHandler(this.ExitToolStripMenuItem_Click);
			// 
			// splitContainer1
			// 
			this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splitContainer1.Location = new System.Drawing.Point(0, 32);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.lvwFiles);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.rtxCode);
			this.splitContainer1.Size = new System.Drawing.Size(978, 712);
			this.splitContainer1.SplitterDistance = 326;
			this.splitContainer1.TabIndex = 1;
			// 
			// ofdExcel
			// 
			this.ofdExcel.Filter = "Excel文件|*.xls;*.xlsx;*.xlsm|所有文件|*.*";
			this.ofdExcel.Multiselect = true;
			// 
			// tmrProgressBar
			// 
			this.tmrProgressBar.Enabled = true;
			// 
			// rtxCode
			// 
			this.rtxCode.Dock = System.Windows.Forms.DockStyle.Fill;
			this.rtxCode.Location = new System.Drawing.Point(0, 0);
			this.rtxCode.Name = "rtxCode";
			this.rtxCode.Size = new System.Drawing.Size(648, 712);
			this.rtxCode.TabIndex = 0;
			this.rtxCode.Text = "";
			// 
			// lvwFiles
			// 
			this.lvwFiles.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lvwFiles.Location = new System.Drawing.Point(0, 0);
			this.lvwFiles.Name = "lvwFiles";
			this.lvwFiles.ShowItemToolTips = true;
			this.lvwFiles.Size = new System.Drawing.Size(326, 712);
			this.lvwFiles.TabIndex = 0;
			this.lvwFiles.UseCompatibleStateImageBehavior = false;
			this.lvwFiles.View = System.Windows.Forms.View.List;
			// 
			// statusStrip1
			// 
			this.statusStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
			this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1});
			this.statusStrip1.Location = new System.Drawing.Point(0, 716);
			this.statusStrip1.Name = "statusStrip1";
			this.statusStrip1.Size = new System.Drawing.Size(978, 28);
			this.statusStrip1.TabIndex = 2;
			this.statusStrip1.Text = "statusStrip1";
			// 
			// toolStripProgressBar1
			// 
			this.toolStripProgressBar1.Name = "toolStripProgressBar1";
			this.toolStripProgressBar1.Size = new System.Drawing.Size(700, 22);
			// 
			// ExceltkGUI
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(978, 744);
			this.Controls.Add(this.statusStrip1);
			this.Controls.Add(this.splitContainer1);
			this.Controls.Add(this.menuStripMain);
			this.MainMenuStrip = this.menuStripMain;
			this.Name = "ExceltkGUI";
			this.Text = "ExceltkGUI";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ExceltkGUI_FormClosing);
			this.menuStripMain.ResumeLayout(false);
			this.menuStripMain.PerformLayout();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			this.splitContainer1.ResumeLayout(false);
			this.statusStrip1.ResumeLayout(false);
			this.statusStrip1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.MenuStrip menuStripMain;
		private System.Windows.Forms.ToolStripMenuItem 文件FToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem OpenFileToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem ExitToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
		private System.Windows.Forms.SplitContainer splitContainer1;
		private System.Windows.Forms.OpenFileDialog ofdExcel;
		private System.Windows.Forms.Timer tmrProgressBar;
		private System.ComponentModel.BackgroundWorker bgwImageRecognition;
		private System.Windows.Forms.ListView lvwFiles;
		private System.Windows.Forms.RichTextBox rtxCode;
		private System.Windows.Forms.StatusStrip statusStrip1;
		private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
	}
}

