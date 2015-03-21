#define _DO_LOG

using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Data;
using System.Timers;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.ServiceProcess;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MySql.Data.MySqlClient;

namespace OneAssistanceSrvc
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
        MySqlConnection _conn, updConn;
        MySqlCommand _cmd;
        string smtpServer;
        int smtpPort;
        string smtpUser;
        string smtpPassword;
        bool enableSSL;
        string notifyTo;

        ArrayList users;

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
			this.ServiceName = "OneAssistanceSrvc";

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
                        ServiceController mService = new ServiceController("OneAssistance");
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
                        ServiceController mService = new ServiceController("OneAssistance");
                        string n = mService.DisplayName;
                    }
                    catch (System.Exception e)
                    {
                        MessageBox.Show(e.Message, "OneAssistanceSrvc /update", MessageBoxButtons.OK);
                        return;
                    }

                    ParamsWin mForm = new ParamsWin();
                    if (mForm.ShowDialog() == DialogResult.OK)
                        MessageBox.Show("Service params successfully updated!", "OneAssistance Service", MessageBoxButtons.OK);
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
                    MessageBox.Show("Version: 1.0.0", "OneAssistance Service", MessageBoxButtons.OK);
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

                this.EventLog.WriteEntry("OneAssistance Service started OK", EventLogEntryType.Information);
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
                EventLog.WriteEntry("Error OneAssistance Service", ie.Message, EventLogEntryType.Error);
				throw(ie);
			}
			m_Timer.Start();
		}

        private bool DoLogin()
        {
            return true;
        }

        private void DoWork()
        {
            if (!OpenDBConnection())
                return;

            _cmd = new MySqlCommand("select count(*) from notifications where prov_id is not null and email_sent_dt is null and (deleted is null or deleted=0)", _conn);
            int cc = Convert.ToInt32(_cmd.ExecuteScalar());
            if (cc == 0)
            {
                DoLogOff();
                return;
            }

            LoadEmails(_cmd);
            SendEmailWithResources("http://www.one-enterprise.net/One-Assistance/CurrentPartners.aspx?id=" + cc,
                "One-Assistance", "adanmn@yahoo.com", users);

            _cmd.CommandText = "update notifications set email_sent_dt=now() where prov_id is not null and email_sent_dt is null and (deleted is null or deleted=0)";
            _cmd.ExecuteNonQuery();

            DoLogOff();
        }

        private void LoadEmails(MySqlCommand cmd)
        {
            users = new ArrayList();
            users.Add(new User("Kytzia Moreno", "kytziam@yahoo.com"));
            users.Add(new User("Disraely Santos", "disraelysn@hotmail.com"));
            users.Add(new User("Adán Mercado", "adan.mercado.n@gmail.com"));
            /*cmd.CommandText = "Select concat(cust_names, ' ', cust_first_name), cust_email from customer";
            MySqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                users.Add(new User(dr.GetString(0), dr.GetString(1)));
            }
            dr.Close();*/
        }

        class User
        {
            public string name;
            public string email;
            public User(string name, string email)
            {
                this.name = name;
                this.email = email;
            }
        }

        private bool OpenDBConnection()//_CHANGE_: change functionality
        {
            bool connOpen = false;
            try
            {
                _conn = new MySqlConnection(
                    "Server=50.31.26.152;Database=biznet;Uid=root;Pwd=vacar00;Pooling=false;Charset=utf8;");
                _conn.Open();
                connOpen = true;
                updConn = new MySqlConnection(
                    "Server=50.31.26.152;Database=biznet;Uid=root;Pwd=vacar00;Pooling=false;Charset=utf8;");
                updConn.Open();
                return true;
            }
            catch (MySqlException oe)
            {
                EventLog.WriteEntry("OneAssistance Service Exception", string.Format("Code:{0} {1}",
                    oe.ErrorCode.ToString(), oe.Message), EventLogEntryType.Information);
                if (connOpen)
                    _conn.Close();
            }

            return false;
        }

        private void DoLogOff()
        {
            updConn.Close();
            _conn.Close();
        }

        private void LoadRegParams()
        {
            RegistryKey mRegKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\OneAssistance\\Monitor");

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

        /// <summary>
        /// Send a web page as an email
        /// </summary>
        /// <param name="url"></param>
        /// <param name="fromName"></param>
        /// <param name="fromEmail"></param>
        /// <param name="toEmail"></param>
        private static void SendEmailWithResources(string url, string fromName, string fromEmail, ArrayList toEmail)
        {
            try
            {
                string html = GetPageHTML(url);
                List<string> images = FindHTMLImages(html);
                string processedPageHTML;
                List<LinkedResource> linkedResources = ProcessEmbeddedHTML(out processedPageHTML, html, images, url);
                SendEmail(linkedResources, processedPageHTML, url, fromName, fromEmail, toEmail);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static void SendEmail(List<LinkedResource> linkedResources, string processedPageHTML, string messageURL, string fromName, string fromEmail, ArrayList toEmail)
        {
            string pwd = "aranchita";
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(fromEmail, fromName);
                mail.To.Add(new MailAddress("adan.mercado.n@gmail.com", "One-Partner"));
                foreach (User item in toEmail)
                {
                    mail.Bcc.Add(new MailAddress(item.email, item.name));
                }
                mail.Subject = "One-Assistance - Nuevos proveedores de productos y servicios";

                string txtBody = "See this email online here: " + messageURL;
                AlternateView plainView = AlternateView.CreateAlternateViewFromString(txtBody, null, "text/plain");

                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(processedPageHTML, null, "text/html");
                foreach (LinkedResource linkedResource in linkedResources)
                    htmlView.LinkedResources.Add(linkedResource);

                mail.AlternateViews.Add(plainView);
                mail.AlternateViews.Add(htmlView);

                SmtpClient client = new SmtpClient("smtp.mail.yahoo.com", 587);
                client.EnableSsl = true;
                client.Credentials = new System.Net.NetworkCredential("adanmn", pwd);
                client.Send(mail);
            }
        }

        /// <summary>
        /// Read the HTML from a URL
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        private static string GetPageHTML(string url)
        {
            string pageHTML = string.Empty;

            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Timeout = 100000;
            using (var stream = request.GetResponse().GetResponseStream())
            {
                using (var reader = new StreamReader(stream))
                {
                    pageHTML = reader.ReadToEnd();
                }
            }

            return pageHTML;
        }

        /// <summary>
        /// Replace the source of all the images with cid:uniqueId, download the image and add to the LinkedResource collection.
        /// </summary>
        /// <param name="pageHTML"></param>
        /// <param name="images"></param>
        /// <param name="baseURL"></param>
        /// <returns></returns>
        private static List<LinkedResource> ProcessEmbeddedHTML(out string processedPageHTML, string pageHTML, List<string> images, string baseURL)
        {
            List<LinkedResource> resources = new List<LinkedResource>();
            foreach (string image in images)
            {
                string imageName = Guid.NewGuid().ToString().Replace("-", string.Empty);
                pageHTML = pageHTML.Replace(image, "cid:" + imageName);
                LinkedResource imagelink = GetLinkedResource(image, baseURL, imageName);
                resources.Add(imagelink);
            }
            processedPageHTML = pageHTML;
            return resources;
        }

        /// <summary>
        /// Create a new LinkedResource from an image stream
        /// </summary>
        /// <param name="imageURL"></param>
        /// <param name="baseURL"></param>
        /// <param name="imageName"></param>
        /// <returns></returns>
        private static LinkedResource GetLinkedResource(string imageURL, string baseURL, string imageName)
        {
            // Turn reletiv URLs to absolute URLs
            Uri imageURI = null;
            if (!Uri.TryCreate(imageURL, UriKind.Absolute, out imageURI))
                Uri.TryCreate(new Uri(baseURL), imageURL, out imageURI);

            if (imageURI == null)
                return null;

            MemoryStream memoryStream = null;
            LinkedResource imagelink = null;
            using (WebClient client = new WebClient())
            {
                byte[] myDataBuffer = client.DownloadData(imageURI);
                memoryStream = new MemoryStream(myDataBuffer);
                imagelink = new LinkedResource(memoryStream, client.ResponseHeaders["Content-Type"]);
            }

            imagelink.ContentId = imageName;
            imagelink.ContentLink = new Uri("cid:" + imageName);
            imagelink.TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
            return imagelink;
        }

        /// <summary>
        /// Get a list of all the images in the HTML
        /// </summary>
        /// <param name="HTMLText"></param>
        /// <returns></returns>
        public static List<string> FindHTMLImages(string HTMLText)
        {
            string anchorPattern = @"(?<=img\s*\S*src\=[\x27\x22])(?<Url>[^\x27\x22]*)(?=[\x27\x22])";
            MatchCollection matches = Regex.Matches(HTMLText, anchorPattern, RegexOptions.IgnorePatternWhitespace | RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled);
            List<string> imageSources = new List<string>();

            foreach (Match m in matches)
            {
                string url = m.Groups["Url"].Value;
                Uri testUri = null;
                if (Uri.TryCreate(url, UriKind.RelativeOrAbsolute, out testUri))
                {
                    if (!imageSources.Exists(s => s == testUri.ToString()))
                        imageSources.Add(testUri.ToString());
                }
            }
            return imageSources;
        }
    }
}
