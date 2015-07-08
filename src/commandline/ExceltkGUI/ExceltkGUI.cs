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
			if (ofdExcel.ShowDialog() == DialogResult.OK)
			{
				//将选择的所有文件的文件名加入列表框，并在 imageFileNames 中保存文件路径
				foreach (string fileName in ofdExcel.FileNames)
				{
					lvwFiles.Items.Add(Path.GetFileName(fileName));
					//此处需要将ShowItemToolTips设为True
					lvwFiles.Items[lvwFiles.Items.Count - 1].ToolTipText = fileName;
				}


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

		private void Convert(string[] files)
		{

		}

		private void bgwConvert_DoWork(object sender, DoWorkEventArgs e)
		{
			//将各参数拆箱
			string tesseractOcrDir = ((List<object>)e.Argument)[0].ToString();
			string outputDir = ((List<object>)e.Argument)[1].ToString();
			List<string> imageFilePaths = (List<string>)((List<object>)e.Argument)[2];
			string langType = ((List<object>)e.Argument)[3].ToString();

			//进度条前期空
			int beforeSpan = 10;
			//进度条后期空
			int afterSpan = 10;

			//先显示一点进度条
			bgwConvert.ReportProgress(beforeSpan);

			for (int i = 0; i < imageFilePaths.Count; i++)
			{
				//取消当前操作
				if (bgwConvert.CancellationPending)
				{
					e.Cancel = true;

					return;
				}


				//获取当前图片路径
				string imageFilePath = imageFilePaths[i];
				//获取图片对应的结果文件的文件路径（默认在图片路径后系统自动加.txt)
				string outputResultFilePath = outputDir + imageFilePath.Substring(imageFilePath.LastIndexOf("\\") + 1);

				//Common.InvokeOcrCommandLine(new Dictionary<string, string>
				//{
				//	{"sourceImagePath", imageFilePath},
				//	{"resultFilePath", outputResultFilePath}
				//});

				CommandLineProcess command = new CommandLineProcess(new Dictionary<string, string>
				{
					{"xlsfile", langType}
				});

				command.Process();


				//报告进度
				bgwConvert.ReportProgress(beforeSpan + (i + 1) * (100 - beforeSpan - afterSpan) / imageFilePaths.Count);
			}

			bgwConvert.ReportProgress(100 - afterSpan);
			//稍微暂停一下，以表现满格前最后一步动作
			Thread.Sleep(1000);

			//进度条满格
			bgwConvert.ReportProgress(100);
			//稍微暂停一下，以表现满格
			Thread.Sleep(500);
		}
	}
}