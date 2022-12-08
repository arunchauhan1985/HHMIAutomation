using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;

namespace eps2math
{
    class executebatch
    {
        string exceptionString = string.Empty;
        public string executeCommand(string command, string argument)
        {
           return ExecuteCommandSync(command, argument);
        }

        public string executeCommandConv(string command, Process proc, bool buttonClicked)
        {
            return ExecuteCommandMain(command, proc, buttonClicked);
        }

        private string ExecuteCommandSync(string command, string argument)
        {
            Process proc = null;
            try
            {
                proc = new Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = command;
                proc.StartInfo.Arguments = argument;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();

                var result = proc.StandardOutput.ReadToEnd();
                Console.WriteLine(result);
                return result;
            }
            catch (Exception)
            {
                return string.Empty;
            }
            finally
            {
                if (proc != null)
                {
                    proc.Close();
                    proc.Dispose();
                }
            }
        }

        private string ExecuteCommandMain(string command, Process proc, bool buttonClicked)
        {
            try
            {
                var procStartInfo = new System.Diagnostics.ProcessStartInfo("cmd", "/c " + command) { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                // The following commands are needed to redirect the standard output.
                proc = new System.Diagnostics.Process { StartInfo = procStartInfo };
                Application.DoEvents();
                proc.Start();
                // Get the output into a string
                var result = proc.StandardOutput.ReadToEnd();
                int status = proc.ExitCode;
                // Display the command output.
                Console.WriteLine(result);

                if (buttonClicked == true)
                {
                    proc.Kill();
                }
            }
            catch (Exception e)
            {
                exceptionString = "Process can't be executed.. ";
                throw;
            }
            finally
            {
                if (proc != null)
                {
                    proc.Close();
                    proc.Dispose();
                }
            } 
            return exceptionString;
        }
    }
}
