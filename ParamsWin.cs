using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Win32;


namespace OneAssistanceSrvc
{
	/// <summary>
	/// Summary description for ParamsWin.
	/// </summary>
	public class ParamsWin : System.Windows.Forms.Form
	{
        private RegValues m_RegValues;
        private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown ddInterval;
        private System.Windows.Forms.TextBox efSystemDSN;
        private System.Windows.Forms.ErrorProvider m_Validator;
        private GroupBox groupBox2;
        private TextBox efSMTPserver;
        private Label label4;
        private GroupBox groupBox5;
        private Label label10;
        private TextBox efDBPassword;
        private TextBox efRetypePwd;
        private Label label3;
        private Label label2;
        private TextBox efDBUser;
        private Label label1;
        private TextBox efSMTPport;
        private TextBox efSMTPpwd;
        private TextBox efSMTPretype;
        private Label label6;
        private Label label8;
        private TextBox efSMTPuser;
        private Label label9;
        private CheckBox cbEnableSSL;
        private Label label11;
        private TextBox efNotifyTo;
        private IContainer components;

		public ParamsWin()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			m_RegValues = new RegValues();

			efSystemDSN.Text = m_RegValues.DBInstance;
			efDBUser.Text = m_RegValues.DBUser;
			efDBPassword.Text = m_RegValues.DBPassword;
            efRetypePwd.Text = m_RegValues.DBPassword;

			ddInterval.Value = Convert.ToDecimal(m_RegValues.ChInterval);

            efSMTPserver.Text = m_RegValues.SMTPServer;
            efSMTPport.Text = m_RegValues.SMTPPort.ToString();
            efSMTPuser.Text = m_RegValues.SMTPUser;
            efSMTPpwd.Text = m_RegValues.SMTPPassword;
            efSMTPretype.Text = m_RegValues.SMTPPassword;
            cbEnableSSL.Checked = m_RegValues.EnableSSL;

            efNotifyTo.Text = m_RegValues.NotifyTo;

			//getRegValues();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		public void setRegValues()
		{
			m_RegValues.DBInstance = efSystemDSN.Text.Trim();
			m_RegValues.DBUser = efDBUser.Text.Trim();
			m_RegValues.DBPassword = efDBPassword.Text.Trim();

			m_RegValues.ChInterval = ddInterval.Value;

            m_RegValues.SMTPServer = efSMTPserver.Text.Trim();
            m_RegValues.SMTPPort = Convert.ToInt32(efSMTPport.Text.Trim());
            m_RegValues.SMTPUser = efSMTPuser.Text.Trim();
            m_RegValues.SMTPPassword = efSMTPpwd.Text.Trim();
            m_RegValues.EnableSSL = cbEnableSSL.Checked;

            m_RegValues.NotifyTo = efNotifyTo.Text.Trim();

            m_RegValues.setRegValues();			
		}


		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.ddInterval = new System.Windows.Forms.NumericUpDown();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.efDBPassword = new System.Windows.Forms.TextBox();
            this.efRetypePwd = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.efDBUser = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.efSystemDSN = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.m_Validator = new System.Windows.Forms.ErrorProvider(this.components);
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbEnableSSL = new System.Windows.Forms.CheckBox();
            this.efSMTPport = new System.Windows.Forms.TextBox();
            this.efSMTPpwd = new System.Windows.Forms.TextBox();
            this.efSMTPretype = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.efSMTPuser = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.efSMTPserver = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.efNotifyTo = new System.Windows.Forms.TextBox();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ddInterval)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_Validator)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(16, 24);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(104, 16);
            this.label5.TabIndex = 4;
            this.label5.Text = "Interval (hours)";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.ddInterval);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Location = new System.Drawing.Point(8, 183);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(312, 56);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Sync Interval";
            // 
            // ddInterval
            // 
            this.ddInterval.Location = new System.Drawing.Point(136, 20);
            this.ddInterval.Maximum = new decimal(new int[] {
            24,
            0,
            0,
            0});
            this.ddInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.ddInterval.Name = "ddInterval";
            this.ddInterval.Size = new System.Drawing.Size(160, 20);
            this.ddInterval.TabIndex = 13;
            this.ddInterval.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(338, 333);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(120, 24);
            this.btnOK.TabIndex = 14;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(530, 334);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(120, 24);
            this.btnCancel.TabIndex = 15;
            this.btnCancel.Text = "Cancel";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.efDBPassword);
            this.groupBox4.Controls.Add(this.efRetypePwd);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Controls.Add(this.label2);
            this.groupBox4.Controls.Add(this.efDBUser);
            this.groupBox4.Controls.Add(this.label1);
            this.groupBox4.Controls.Add(this.efSystemDSN);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Location = new System.Drawing.Point(8, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(312, 154);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "MySQL ODBC";
            // 
            // efDBPassword
            // 
            this.efDBPassword.Location = new System.Drawing.Point(136, 81);
            this.efDBPassword.Name = "efDBPassword";
            this.efDBPassword.PasswordChar = '*';
            this.efDBPassword.Size = new System.Drawing.Size(160, 20);
            this.efDBPassword.TabIndex = 15;
            // 
            // efRetypePwd
            // 
            this.efRetypePwd.Location = new System.Drawing.Point(136, 116);
            this.efRetypePwd.Name = "efRetypePwd";
            this.efRetypePwd.PasswordChar = '*';
            this.efRetypePwd.Size = new System.Drawing.Size(160, 20);
            this.efRetypePwd.TabIndex = 16;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(16, 116);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 16);
            this.label3.TabIndex = 13;
            this.label3.Text = "Retype Password";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(16, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 16);
            this.label2.TabIndex = 12;
            this.label2.Text = "DB Password";
            // 
            // efDBUser
            // 
            this.efDBUser.Location = new System.Drawing.Point(136, 49);
            this.efDBUser.Name = "efDBUser";
            this.efDBUser.Size = new System.Drawing.Size(160, 20);
            this.efDBUser.TabIndex = 14;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(104, 16);
            this.label1.TabIndex = 11;
            this.label1.Text = "DB User";
            // 
            // efSystemDSN
            // 
            this.efSystemDSN.CausesValidation = false;
            this.efSystemDSN.Location = new System.Drawing.Point(136, 16);
            this.efSystemDSN.Name = "efSystemDSN";
            this.efSystemDSN.Size = new System.Drawing.Size(160, 20);
            this.efSystemDSN.TabIndex = 2;
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(16, 20);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(104, 16);
            this.label7.TabIndex = 0;
            this.label7.Text = "System DSN";
            // 
            // m_Validator
            // 
            this.m_Validator.ContainerControl = this;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbEnableSSL);
            this.groupBox2.Controls.Add(this.efSMTPport);
            this.groupBox2.Controls.Add(this.efSMTPpwd);
            this.groupBox2.Controls.Add(this.efSMTPretype);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.efSMTPuser);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.efSMTPserver);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Location = new System.Drawing.Point(338, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(312, 227);
            this.groupBox2.TabIndex = 16;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "SMTP Server";
            // 
            // cbEnableSSL
            // 
            this.cbEnableSSL.AutoSize = true;
            this.cbEnableSSL.Location = new System.Drawing.Point(136, 180);
            this.cbEnableSSL.Name = "cbEnableSSL";
            this.cbEnableSSL.Size = new System.Drawing.Size(82, 17);
            this.cbEnableSSL.TabIndex = 20;
            this.cbEnableSSL.Text = "Enable SSL";
            this.cbEnableSSL.UseVisualStyleBackColor = true;
            // 
            // efSMTPport
            // 
            this.efSMTPport.CausesValidation = false;
            this.efSMTPport.Location = new System.Drawing.Point(136, 49);
            this.efSMTPport.Name = "efSMTPport";
            this.efSMTPport.Size = new System.Drawing.Size(160, 20);
            this.efSMTPport.TabIndex = 19;
            // 
            // efSMTPpwd
            // 
            this.efSMTPpwd.Location = new System.Drawing.Point(136, 113);
            this.efSMTPpwd.Name = "efSMTPpwd";
            this.efSMTPpwd.PasswordChar = '*';
            this.efSMTPpwd.Size = new System.Drawing.Size(160, 20);
            this.efSMTPpwd.TabIndex = 16;
            // 
            // efSMTPretype
            // 
            this.efSMTPretype.Location = new System.Drawing.Point(136, 145);
            this.efSMTPretype.Name = "efSMTPretype";
            this.efSMTPretype.PasswordChar = '*';
            this.efSMTPretype.Size = new System.Drawing.Size(160, 20);
            this.efSMTPretype.TabIndex = 17;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(16, 149);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(104, 16);
            this.label6.TabIndex = 13;
            this.label6.Text = "Retype Password";
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(16, 116);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(104, 16);
            this.label8.TabIndex = 14;
            this.label8.Text = "Password";
            // 
            // efSMTPuser
            // 
            this.efSMTPuser.Location = new System.Drawing.Point(136, 81);
            this.efSMTPuser.Name = "efSMTPuser";
            this.efSMTPuser.Size = new System.Drawing.Size(160, 20);
            this.efSMTPuser.TabIndex = 15;
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(16, 85);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(104, 16);
            this.label9.TabIndex = 12;
            this.label9.Text = "User";
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(16, 49);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(104, 16);
            this.label10.TabIndex = 3;
            this.label10.Text = "Port";
            // 
            // efSMTPserver
            // 
            this.efSMTPserver.CausesValidation = false;
            this.efSMTPserver.Location = new System.Drawing.Point(136, 16);
            this.efSMTPserver.Name = "efSMTPserver";
            this.efSMTPserver.Size = new System.Drawing.Size(160, 20);
            this.efSMTPserver.TabIndex = 2;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(16, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(104, 16);
            this.label4.TabIndex = 0;
            this.label4.Text = "Domain Name";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label11);
            this.groupBox5.Controls.Add(this.efNotifyTo);
            this.groupBox5.Location = new System.Drawing.Point(8, 257);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(642, 56);
            this.groupBox5.TabIndex = 17;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Notify To";
            // 
            // label11
            // 
            this.label11.Location = new System.Drawing.Point(16, 22);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(104, 16);
            this.label11.TabIndex = 18;
            this.label11.Text = "Email(s)";
            // 
            // efNotifyTo
            // 
            this.efNotifyTo.Location = new System.Drawing.Point(136, 19);
            this.efNotifyTo.Name = "efNotifyTo";
            this.efNotifyTo.Size = new System.Drawing.Size(490, 20);
            this.efNotifyTo.TabIndex = 17;
            // 
            // ParamsWin
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(681, 370);
            this.ControlBox = false;
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.groupBox3);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ParamsWin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "OneAssistance - Setup";
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ddInterval)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_Validator)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

        private void btnOK_Click(object sender, System.EventArgs e)
		{
			bool valid = true;
            this.efSystemDSN.Text = this.efSystemDSN.Text.Trim();
			if (efSystemDSN.Text.Length==0 )
			{
				MessageBox.Show("You must provide an ODBC System DSN","Validating entries",MessageBoxButtons.OK);
				//efNetSrvcName.Select(0,efNetSrvcName.Text.Length);
				efSystemDSN.Focus();
				valid = false;
				return;
			}
			/*if (efDBUser.Text.Trim().Length==0)
			{
				MessageBox.Show("You must provide the DB User","Validating entries",MessageBoxButtons.OK);
				//efDBUser.Select(0,efDBUser.Text.Length);
				efDBUser.Focus();
				valid = false;
				return;
			}
			if (efDBPassword.Text.Trim().Length==0)
			{
				MessageBox.Show("You must provide the DB Password","Validating entries",MessageBoxButtons.OK);
				//efDBPassword.Select(0,efDBPassword.Text.Length);
				efDBPassword.Focus();
				valid = false;
				return;
			}
			if (this.efRetypePwd.Text.Trim().Length==0)
			{
				MessageBox.Show("You must retype a valid DB Password","Validating entries",MessageBoxButtons.OK);
				//efRetypePwd.Select(0,efRetypePwd.Text.Length);
				efRetypePwd.Focus();
				valid = false;
			}*/

            this.efDBUser.Text = this.efDBUser.Text.Trim();
            this.efDBPassword.Text = this.efDBPassword.Text.Trim();
            this.efRetypePwd.Text = this.efRetypePwd.Text.Trim();
            if (efDBPassword.Text.Length > 0 || efRetypePwd.Text.Length > 0)
            {
                if (this.efDBPassword.Text.CompareTo(this.efRetypePwd.Text) != 0)
                {
                    MessageBox.Show("Both password entries must be equal", "Validating entries", MessageBoxButtons.OK);
                    if (efDBPassword.Text.Length > 0)
                        efDBPassword.Select(0, efDBPassword.Text.Length);
                    efDBPassword.Focus();
                    valid = false;
                    return;
                }
                if (efDBUser.Text.Length == 0)
                {
                    MessageBox.Show("You must provide the DB User", "Validating entries", MessageBoxButtons.OK);
                    efDBUser.Focus();
                    valid = false;
                    return;
                }
            }

            /////////////////////////////////////////////////////////////////

            this.efSMTPserver.Text = this.efSMTPserver.Text.Trim();
            if (efSMTPserver.Text.Length == 0)
            {
                MessageBox.Show("You must provide the SMTP Server domain name or IP address", "Validating entries", MessageBoxButtons.OK);
                efSMTPserver.Focus();
                valid = false;
                return;
            }
            this.efSMTPport.Text = this.efSMTPport.Text.Trim();
            if (efSMTPport.Text.Length == 0)
            {
                MessageBox.Show("You must provide the SMTP Port", "Validating entries", MessageBoxButtons.OK);
                efSMTPport.Focus();
                valid = false;
                return;
            }
            this.efSMTPuser.Text = this.efSMTPuser.Text.Trim();
            if (efSMTPuser.Text.Length == 0)
            {
                MessageBox.Show("You must provide the SMTP User", "Validating entries", MessageBoxButtons.OK);
                efSMTPuser.Focus();
                valid = false;
                return;
            }
            this.efSMTPpwd.Text = this.efSMTPpwd.Text.Trim();
            if (efSMTPpwd.Text.Length == 0)
            {
                MessageBox.Show("You must provide the SMTP Password", "Validating entries", MessageBoxButtons.OK);
                efSMTPpwd.Focus();
                valid = false;
                return;
            }
            this.efSMTPretype.Text = this.efSMTPretype.Text.Trim();
            if (this.efSMTPretype.Text.Length == 0)
            {
                MessageBox.Show("You must retype a valid SMTP Password", "Validating entries", MessageBoxButtons.OK);
                efSMTPretype.Focus();
                valid = false;
            }

            /////////////////////////////////////////////////////////////////////

            if (this.efSMTPpwd.Text.CompareTo(this.efSMTPretype.Text) != 0)
            {
                MessageBox.Show("SMTP: Both password entries must be equal", "Validating entries", MessageBoxButtons.OK);
                efSMTPpwd.Select(0, efSMTPpwd.Text.Length);
                efSMTPpwd.Focus();
                valid = false;
                return;
            }

            this.efNotifyTo.Text = this.efNotifyTo.Text.Trim();
            if (efNotifyTo.Text.Length == 0)
            {
                MessageBox.Show("You must provide at least an email account for notificacions", "Validating entries", MessageBoxButtons.OK);
                efNotifyTo.Focus();
                valid = false;
            }

            if (valid)
			{
				setRegValues();
				this.DialogResult = DialogResult.OK;
			}
		}
	}
}
