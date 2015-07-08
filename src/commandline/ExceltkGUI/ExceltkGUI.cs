using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
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
	}
}