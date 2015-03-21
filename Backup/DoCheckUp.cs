#define _DO_LOG

using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Data;
using System.Data.OleDb;
using System.Timers;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.ServiceProcess;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Text;

namespace PTMSrvc
{
	public class DoCheckUp : System.ServiceProcess.ServiceBase
	{
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public System.Timers.Timer m_Timer;

		private RegValues m_RegValues;

		public double m_TimeStep = 600000;

        private string sSecurityKey;
        string DBInstance;
        string DBUser;
        string DBPassword;
        OleDbConnection _conn, updConn;
        OleDbCommand _cmd;
        string smtpServer;
        int smtpPort;
        string smtpUser;
        string smtpPassword;
        bool enableSSL;
        string notifyTo;

		public DoCheckUp()
		{
			// This call is required by the Windows.Forms Component Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitComponent call
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			// 
			// CheckUp
			// 
			this.ServiceName = "PTM";

            LoadRegParams();
		}
		#endregion

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing ) 
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		// The main entry point for the process
		static void Main(string[] args)
		{
			string opt = null;
			// check for arguments
			if (args.Length > 0)
			{
				opt = args[0];
				if (opt != null && opt.ToLower() == "/install")
				{
					ParamsWin mForm = new ParamsWin();
					if (mForm.ShowDialog() == DialogResult.Cancel)
						return;

					TransactedInstaller ti = new TransactedInstaller();
					ProjectInstaller pi = new ProjectInstaller();
					ti.Installers.Add(pi);
					
					String path = String.Format("/assemblypath={0}",
						Assembly.GetExecutingAssembly().Location);
					String[] cmdline = {path};
					InstallContext ctx = new InstallContext("", cmdline);

					ti.Context = ctx;
					try
					{
                        ServiceController mService = new ServiceController("_CHANGE_");
						if (mService.DisplayName!=string.Empty)
						{
							if (mService.CanStop)
							{
								mService.Stop();
								mService.WaitForStatus(ServiceControllerStatus.Stopped);
							}
							ti.Uninstall(null);
						}
					}
					catch (System.Exception)
					{
					}

					ti.Install(new Hashtable());
                }
                else if (opt != null && opt.ToLower() == "/update")
                {
                    try
                    {
                        ServiceController mService = new ServiceController("_CHANGE_");
                        string n = mService.DisplayName;
                    }
                    catch (System.Exception e)
                    {
                        MessageBox.Show(e.Message, "_CHANGE_Srvc /update", MessageBoxButtons.OK);
                        return;
                    }

                    ParamsWin mForm = new ParamsWin();
                    if (mForm.ShowDialog() == DialogResult.OK)
                        MessageBox.Show("Service params successfully updated!", "_CHANGE_ Service", MessageBoxButtons.OK);
                }
				else if (opt != null && opt.ToLower() == "/uninstall")
				{
					TransactedInstaller ti = new TransactedInstaller();
					ProjectInstaller mi = new ProjectInstaller();
					ti.Installers.Add(mi);
					String path = String.Format("/assemblypath={0}",
						Assembly.GetExecutingAssembly().Location);
					String[] cmdline = {path};
					InstallContext ctx = new InstallContext("", cmdline);
					ti.Context = ctx;
                    try
					{
					    ti.Uninstall(null);
                    }
                    catch (System.Exception)
                    {
                    }
					RegValues lRegValues = new RegValues();
					lRegValues.clearRegValues();
				}
                else if (opt != null && opt.ToLower() == "/version")
                {
                    MessageBox.Show("Version: 1.0.0", "_CHANGE_ Service", MessageBoxButtons.OK);
                }
			}

			if (opt == null) // e.g. ,nothing on the command line
			{
#if ( !DEBUG )
				System.ServiceProcess.ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] {new DoCheckUp()};
                ServiceBase.Run(ServicesToRun);
#else
				// debug code: allows the process to run as a non-service
				// will kick off the service start point, but never kill it
				// shut down the debugger to exit
                DoIntegration service = new DoIntegration();
                service.OnStart(null);
                Thread.Sleep(Timeout.Infinite);
#endif 
			}
		}

		/// <summary>
		/// Set things in motion so your service can do its work.
		/// </summary>
		protected override void OnStart(string[] args)
		{
			try
			{
				m_RegValues = new RegValues();
                //_CHANGE_
				m_TimeStep = Convert.ToDouble(m_RegValues.ChInterval) * 60000; //Minutes
                //m_TimeStep = (Convert.ToDouble(m_RegValues.ChInterval) * 60) * 60000; //Hours

				m_Timer = new System.Timers.Timer();
				m_Timer.Elapsed += new ElapsedEventHandler(myTimer_Elapsed);
				m_Timer.AutoReset = true;
				m_Timer.Interval = m_TimeStep;
				m_Timer.Enabled = false;
				//			nextInc = this.getIncNearTime(m_Start);//new DateTime(1971,1,1);
				//			nextFull = this.getFullNearTime(m_Start);//new DateTime(1971,1,1);
				this.m_Timer.Start();

                this.EventLog.WriteEntry("_CHANGE_ Service started OK", EventLogEntryType.Information);
			}
			catch(System.Exception e)
			{
				System.Exception ie = e;
				while(ie.InnerException!=null)
					ie = ie.InnerException;
				throw(ie);
			}
		}
 
		/// <summary>
		/// Stop this service.
		/// </summary>
		protected override void OnStop()
		{
			// TODO: Add code here to perform any tear-down necessary to stop your service.
		}

		private void myTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
		{
			m_Timer.Stop();
			try
			{
				DoWork();
			}
			catch(System.Exception ex)
			{
				System.Exception ie = ex;
				while(ie.InnerException!=null)
					ie = ie.InnerException;
                EventLog.WriteEntry("Error _CHANGE_ Service", ie.Message, EventLogEntryType.Error);
				throw(ie);
			}
			m_Timer.Start();
		}

        private bool DoLogin()
        {
            return false;
        }

        private void DoLogOff()
        {
            updConn.Close();
            _conn.Close();
        }

        private void DoWork()
        {
            if (!OpenDBConnection())
                return;

            //_CHANGE_: add check-up process here

            DoLogOff();
        }

        private bool OpenDBConnection()//_CHANGE_: change functionality
        {
            bool connOpen = false;
            try
            {
                _conn = new OleDbConnection("Provider=MSDataShape.1;Persist Security Info=False;" +
                    //"User ID=" + DBUser + ";Password=" + DBPassword + ";" + 
                    "Data Source=" + DBInstance + ";packet size=4096");
                _conn.Open();
                connOpen = true;
                updConn = new OleDbConnection("Provider=MSDataShape.1;Persist Security Info=False;" +
                    //"User ID=" + DBUser + ";Password=" + DBPassword + ";" +
                    "Data Source=" + DBInstance + ";packet size=4096");
                updConn.Open();
                return true;
            }
            catch (OleDbException oe)
            {
                EventLog.WriteEntry("_CHANGE_ Service Exception", string.Format("Code:{0} {1}",
                    oe.ErrorCode.ToString(), oe.Message), EventLogEntryType.Information);
                if (connOpen)
                    _conn.Close();
            }

            return false;
        }

        private void LoadRegParams()
        {
            RegistryKey mRegKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\_CHANGE_\\Alerts");

            DBInstance = getRegValue(mRegKey, "SystemDSN").ToString();

            DBUser = getRegValue(mRegKey, "DBUser").ToString();
            DBUser = Encoding.Unicode.GetString(Convert.FromBase64String(DBUser));
            DBPassword = getRegValue(mRegKey, "DBPassword").ToString();
            DBPassword = Encoding.Unicode.GetString(Convert.FromBase64String(DBPassword));

            smtpServer = getRegValue(mRegKey, "SMTPServer").ToString();
            smtpPort = Convert.ToInt32(getRegValue(mRegKey, "SMTPPort"));
            smtpUser = getRegValue(mRegKey, "SMTPUser").ToString();
            smtpUser = Encoding.Unicode.GetString(Convert.FromBase64String(smtpUser));
            smtpPassword = getRegValue(mRegKey, "SMTPPassword").ToString();
            smtpPassword = Encoding.Unicode.GetString(Convert.FromBase64String(smtpPassword));
            enableSSL = Boolean.Parse(getRegValue(mRegKey, "EnableSSL"));

            notifyTo = getRegValue(mRegKey, "NotifyTo").ToString();

            mRegKey.Close();
        }

        private string getRegValue(RegistryKey mRegKey, string keyName)
        {
            string result = null;

            if (mRegKey.GetValue(keyName) != null)
            {
                result = mRegKey.GetValue(keyName).ToString();
            }

            return result;
        }

	}
}
