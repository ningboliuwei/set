using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExceltkGUI
{
	public partial class ExceltkGUI : Form
	{
		public ExceltkGUI()
		{
			InitializeComponent();
		}

		private void OpenFileToolStripMenuItem_Click(object sender, EventArgs e)
		{
			List<string> xlsFiles = new List<string>();

			if (ofdExcel.ShowDialog() == DialogResult.OK)
			{
				//将选择的所有文件的文件名加入列表框合并在 imageFileNames 中保存文件路径
				foreach (string fileName in ofdExcel.FileNames)
				{
					lvwFiles.Items.Add(Path.GetFileName(fileName));
					//此处需要将ShowItemToolTips设为True
					lvwFiles.Items[lvwFiles.Items.Count - 1].ToolTipText = fileName;
					xlsFiles.Add(fileName);
				}

				bgwConvert.RunWorkerAsync(xlsFiles);
			}
		}

		private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
		{
			Application.Exit();
		}

		private void ExceltkGUI_FormClosing(object sender, FormClosingEventArgs e)
		{
			//关闭窗体
			if (MessageBox.Show(this, "确定退出程序吗?", "问题", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2) == DialogResult.No)
			{
				e.Cancel = true;
			}
		}


		private void bgwConvert_DoWork(object sender, DoWorkEventArgs e)
		{
			//将各参数拆箱
			List<string> xlsFiles = (List<string>)e.Argument;

			//进度条前期空 
			//int beforeSpan = 10;
			//进度条后期空
			//int afterSpan = 10;

			//先显示一点进度条
			//bgwConvert.ReportProgress(beforeSpan);

			for (int i = 0; i < xlsFiles.Count; i++)
			{
				//取消当前操作
				if (bgwConvert.CancellationPending)
				{
					e.Cancel = true;

					return;
				}

				CommandLineProcess command = new CommandLineProcess(new Dictionary<string, string> { { "xlsFile", xlsFiles[i] } });
				command.Process();

				//报告进度
				//bgwConvert.ReportProgress(beforeSpan + (i + 1) * (100 - beforeSpan - afterSpan) / xlsFiles.Count);
			}

			//bgwConvert.ReportProgress(100 - afterSpan);
			//稍微暂停一下，以表现满格前最后一步动作
			//Thread.Sleep(1000);

			//进度条满格
			//bgwConvert.ReportProgress(100);
			//稍微暂停一下，以表现满格
			//Thread.Sleep(500);
		}

		private void lvwFiles_SelectedIndexChanged(object sender, EventArgs e)
		{
			//MessageBox.Show(lvwFiles.SelectedItems);
			if (lvwFiles.SelectedItems.Count != 0)
			{
				string xlsFile = lvwFiles.Items[lvwFiles.SelectedIndices[0]].ToolTipText;
				string markDownFileName = Path.GetFileNameWithoutExtension(xlsFile);
				string markDownFilePath = Path.Combine(Path.GetDirectoryName(xlsFile), markDownFileName + "Rank.md");

				ShowMarkDownFile(markDownFilePath);
			}
		}

		private void ShowMarkDownFile(string markDownPath)
		{
			try
			{
				rtxCode.Text = File.ReadAllText(markDownPath);
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
		}
	}
}