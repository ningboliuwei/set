using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExceltkGUI
{
	class CommandLineProcess
	{
		private Dictionary<string, string> _args;

		public CommandLineProcess(Dictionary<string, string> args)
		{
			_args = args;
		}

		public void Process()
		{
			try
			{
				using (Process process = new Process())
				{
					process.StartInfo.FileName = Path.Combine(Application.StartupPath, "exceltk.exe");
					process.StartInfo.Arguments = string.Format("-t md -xls {0}", _args["xlsFile"]);
					process.StartInfo.UseShellExecute = false;
					process.StartInfo.CreateNoWindow = true;
					process.StartInfo.RedirectStandardOutput = true;
					process.StartInfo.RedirectStandardError = true;
					process.StartInfo.WorkingDirectory = Application.StartupPath;
					process.Start();
					process.WaitForExit();
				}
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message);
			}
		}
	}
}
