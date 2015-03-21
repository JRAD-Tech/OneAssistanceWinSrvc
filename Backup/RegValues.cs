using System;
using Microsoft.Win32;
//using Blowfish_NET;
using System.Text;

namespace PTMSrvc
{
	/// <summary>
	/// Summary description for RegValues.
	/// </summary>
	public class RegValues
	{
		private string m_DBInstance;
		private string m_DBUser;
		private string m_DBPassword;
		private decimal m_ChInterval;
        private string m_SMTPserver;
        private int m_SMTPport;
        private string m_SMTPuser;
        private string m_SMTPpwd;
        private bool m_EnableSSL;
        private string m_NotifyTo;

		public string DBInstance
		{
			get 
			{
				return m_DBInstance;
			}
			set 
			{
				m_DBInstance = value;
			}
		}

        public string DBUser
		{
			get 
			{
				return m_DBUser;
			}
			set 
			{
				m_DBUser = value;
			}
		}
		public string DBPassword
		{
			get 
			{
				return m_DBPassword;
			}
			set 
			{
				m_DBPassword = value;
			}
		}

        public string SMTPServer
        {
            get
            {
                return m_SMTPserver;
            }
            set
            {
                m_SMTPserver = value;
            }
        }

        public int SMTPPort
        {
            get
            {
                return m_SMTPport;
            }
            set
            {
                m_SMTPport = value;
            }
        }

        public string SMTPUser
        {
            get
            {
                return m_SMTPuser;
            }
            set
            {
                m_SMTPuser = value;
            }
        }

        public string SMTPPassword
        {
            get
            {
                return m_SMTPpwd;
            }
            set
            {
                m_SMTPpwd = value;
            }
        }

        public bool EnableSSL
        {
            get
            {
                return m_EnableSSL;
            }
            set
            {
                m_EnableSSL = value;
            }
        }

        public string NotifyTo
        {
            get
            {
                return m_NotifyTo;
            }
            set
            {
                m_NotifyTo = value;
            }
        }

        public decimal ChInterval
		{
			get 
			{
				return m_ChInterval;
			}
			set 
			{
				m_ChInterval = value;
			}
		}

		public RegValues()
		{
			getRegValues();
		}

		public void getRegValues()
		{

            RegistryKey mRegKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\_CHANGE_\\Alerts");

            if (getRegValue(mRegKey, "SystemDSN") == null)
			{
                setRegValue(mRegKey, "SystemDSN", "PTM");//_CHANGE_
				m_DBInstance = "PTM";
			}
			else
                m_DBInstance = getRegValue(mRegKey, "SystemDSN").ToString();

            if (getRegValue(mRegKey, "DBUser") == null)
            {
                setRegValue(mRegKey, "DBUser", ""
                    /*Convert.ToBase64String(Encoding.Unicode.GetBytes(""))*/);
                m_DBUser = "";
            }
            else
            {
                m_DBUser = getRegValue(mRegKey, "DBUser").ToString();
                if (m_DBUser.Length > 0)
                    m_DBUser = Encoding.Unicode.GetString(Convert.FromBase64String(m_DBUser));
            }

            if (getRegValue(mRegKey, "DBPassword") == null)
            {
                setRegValue(mRegKey, "DBPassword", ""
                    /*Convert.ToBase64String(Encoding.Unicode.GetBytes(""))*/);
                m_DBPassword = "";
            }
            else
            {
                m_DBPassword = getRegValue(mRegKey, "DBPassword").ToString();
                if (m_DBPassword.Length > 0)
                    m_DBPassword = Encoding.Unicode.GetString(Convert.FromBase64String(m_DBPassword));
            }

			if(mRegKey.GetValue("ChInterval")==null)
			{
				mRegKey.SetValue("ChInterval","1");
				m_ChInterval = 1;
			}
			else
				m_ChInterval = Convert.ToDecimal(mRegKey.GetValue("ChInterval").ToString());

            if (getRegValue(mRegKey, "SMTPServer") == null)
            {
                setRegValue(mRegKey, "SMTPServer", "");
                m_SMTPserver = "";
            }
            else
                m_SMTPserver = getRegValue(mRegKey, "SMTPServer").ToString();

            if (mRegKey.GetValue("SMTPPort") == null)
            {
                mRegKey.SetValue("SMTPPort", "465");
                m_SMTPport = 465;
            }
            else
                m_SMTPport = Convert.ToInt32(mRegKey.GetValue("SMTPPort").ToString());

            if (getRegValue(mRegKey, "SMTPUser") == null)
            {
                setRegValue(mRegKey, "SMTPUser", ""
                    /*Convert.ToBase64String(Encoding.Unicode.GetBytes("SMTP User"))*/);
                m_SMTPuser = "";
            }
            else
            {
                m_SMTPuser = getRegValue(mRegKey, "SMTPUser").ToString();
                m_SMTPuser = Encoding.Unicode.GetString(Convert.FromBase64String(m_SMTPuser));
            }

            if (getRegValue(mRegKey, "SMTPPassword") == null)
            {
                setRegValue(mRegKey, "SMTPPassword", ""
                    /*Convert.ToBase64String(Encoding.Unicode.GetBytes("SMTP Password"))*/);
                m_SMTPpwd = "";
            }
            else
            {
                m_SMTPpwd = getRegValue(mRegKey, "SMTPPassword").ToString();
                m_SMTPpwd = Encoding.Unicode.GetString(Convert.FromBase64String(m_SMTPpwd));
            }

            if (getRegValue(mRegKey, "EnableSSL") == null)
            {
                setRegValue(mRegKey, "EnableSSL", "False");
                m_EnableSSL = false;
            }
            else
                m_EnableSSL = Boolean.Parse(getRegValue(mRegKey, "EnableSSL"));

            if (getRegValue(mRegKey, "NotifyTo") == null)
            {
                setRegValue(mRegKey, "NotifyTo", "");
                m_NotifyTo = "";
            }
            else
                m_NotifyTo = getRegValue(mRegKey, "NotifyTo").ToString();

            mRegKey.Close();
		}

		public void setRegValues()
		{

            RegistryKey mRegKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\_CHANGE_\\Alerts");

            setRegValue(mRegKey, "SystemDSN", m_DBInstance);

			setRegValue(mRegKey,"DBUser",
                Convert.ToBase64String(Encoding.Unicode.GetBytes(m_DBUser.Trim())));
			setRegValue(mRegKey,"DBPassword",
                Convert.ToBase64String(Encoding.Unicode.GetBytes(m_DBPassword.Trim())));

			mRegKey.SetValue("ChInterval",m_ChInterval.ToString());

            setRegValue(mRegKey, "SMTPServer", m_SMTPserver);
            setRegValue(mRegKey, "SMTPPort", m_SMTPport.ToString());
            setRegValue(mRegKey, "SMTPUser",
                Convert.ToBase64String(Encoding.Unicode.GetBytes(m_SMTPuser.Trim())));
            setRegValue(mRegKey, "SMTPPassword",
                Convert.ToBase64String(Encoding.Unicode.GetBytes(m_SMTPpwd.Trim())));
            mRegKey.SetValue("EnableSSL", m_EnableSSL.ToString());

            setRegValue(mRegKey, "NotifyTo", m_NotifyTo);

            mRegKey.Close();

		}

		private void setRegValue(RegistryKey mRegKey, string keyName, string keyValue)
		{
			mRegKey.SetValue(keyName, keyValue/*keyEncripted*/);
		}

		private string getRegValue(RegistryKey mRegKey, string keyName)
		{
			string result = null;

			if ( mRegKey.GetValue(keyName)!=null)
			{
				result = mRegKey.GetValue(keyName).ToString();
			}

			return result;
		}

		public void clearRegValues()
		{
            RegistryKey mRegKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\_CHANGE_");
            mRegKey.DeleteSubKey("Alerts");
			mRegKey.Close();
		}

		public bool checkKey()
		{
			bool result = true;
            RegistryKey mRegKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\_CHANGE_");
            if (mRegKey.OpenSubKey("Alerts") == null)
			{
				result = false;
			}
			mRegKey.Close();
			return result;
		}
	}
}
