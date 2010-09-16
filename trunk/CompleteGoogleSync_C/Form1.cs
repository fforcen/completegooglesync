/*
    Copyright (C) 2010 by Fernando Forcén López

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using GoogleSynchronizer;
using System.Threading;

namespace GCompleteSync
{
    public partial class frmPrincipal : Form
    {
        //delegate void SetTextCallback(string text);
        private Synchronizer sync = new Synchronizer();
        //Thread syncThread;

        public frmPrincipal()
        {
            InitializeComponent();
        }

        /*private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.lbStatus.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.lbStatus.Text = text;
            }
        }*/

        private void Form1_Load(object sender, EventArgs e)
        {
            sync.ChangeStatusEvent += ChangeStatusEventHandler;

            lbStatus.Text = "";

            notifyIcon1.Visible = false;
            notifyIcon1.Text = Application.ProductName;

            //cmbSynFrequency.Items.Add("1 Minuto");
            cmbSynFrequency.Items.Add("5 Minutos");
            cmbSynFrequency.Items.Add("15 Minutos");
            cmbSynFrequency.Items.Add("30 Minutos");

            cmbSynFrequency.SelectedIndex = 0;
            
            cmbSynPriority.DataSource = Enum.GetNames(typeof(SyncPriority));

            chkCustomCategory.Checked = sync.UseCustomCatergory;

            if (sync.StoreGoogleAccount != "")
            {
                txtUsername.Text = sync.StoreGoogleAccount;
            }
            if (sync.StoreGoogleAccountPassword != "")
            {
                txtPassword.Text = sync.StoreGoogleAccountPassword;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Synchronize();
        }

        private void Synchronize()
        {
            if (txtUsername.Text != "" && txtPassword.Text != "")
            {
                //if (syncThread!=null && !syncThread.IsAlive)
                //{
                    timer1.Enabled = false;
                    sync.GoogleAccountUsername = txtUsername.Text;
                    sync.GoogleAccountPassword = txtPassword.Text;
                    sync.SyncPriority = (SyncPriority)cmbSynPriority.SelectedIndex;
                    sync.UseCustomCatergory = chkCustomCategory.Checked;
                    sync.Synchronize();
                    timer1.Enabled = true;
                    //syncThread = new Thread(sync.Synchronize);
                    //syncThread.IsBackground = true;
                    //syncThread.Start();
                    //syncThread.Join();
                //}
            }
        }

        void ChangeStatusEventHandler(SynchronizationStatus newStatus,String message)
        {
            if (newStatus == SynchronizationStatus.Synchronized)
            {
                lbStatus.Text = "Última sincronización: " + DateTime.Now;
                //SetText("Última sincronización: " + DateTime.Now);
                Application.DoEvents();
                notifyIcon1.Text = Application.ProductName + ": " + DateTime.Now;
            }
            else if (newStatus==SynchronizationStatus.SynchronizingContacts)
            {
                lbStatus.Text = "Última sincronización: Sincronizando contactos...";
                //SetText("Última sincronización: Sincronizando contactos...");
                Application.DoEvents();
            }
            else if (newStatus == SynchronizationStatus.SynchronizingCalendar)
            {
                lbStatus.Text = "Última sincronización: Sincronizando calendario...";
                //SetText("Última sincronización: Sincronizando calendario...");
                Application.DoEvents();
            }
            else if (newStatus == SynchronizationStatus.SynchronizationError)
            {
                lbStatus.Text = message;
                Application.DoEvents();
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ShowMainWindow();
        }

        private void ShowMainWindow()
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }

        private void frmPrincipal_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                notifyIcon1.Visible = true;
                Hide();
            }
        }

        private void cmbSynFrequency_SelectedIndexChanged(object sender, EventArgs e)
        {
            String[] minutos;
            if (cmbSynFrequency.SelectedText != "")
            {
                minutos = cmbSynFrequency.SelectedText.Split(' ');
                timer1.Interval = Int32.Parse(minutos[0]) * 60000;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Synchronize();
        }

        private void exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void showMainWindow_Click(object sender, EventArgs e)
        {
            ShowMainWindow();
        }
    }
}