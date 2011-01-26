/*
 * Created by SharpDevelop.
 * User: oferfrid
 * Date: 22/03/2009
 * Time: 13:16
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using CoolLedControl;

namespace LedTest
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		LedControl LC;
		
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		void MainFormLoad(object sender, EventArgs e)
		{
			LC = new LedControl();
			
		}
		
		void MainFormFormClosing(object sender, FormClosingEventArgs e)
		{
			LC.closeSock();
		}
		
		void ConnectDisconnectClick(object sender, EventArgs e)
		{
			if (!ConnectDisconnect.Checked)
			{
				LC.closeSock();
				ConnectDisconnect.Text = "Connect";
				ConnectDisconnect.Checked = false;
			}
			else
			{
				LC.openSock();
				ConnectDisconnect.Text = "Disconnect";
				ConnectDisconnect.Checked  = true;
			}
		}
		
		void LedCheckedChanged(object sender, EventArgs e)
		{
			int ind = GetControlID(sender);
			CheckBox CB = (CheckBox)sender;
			if (CB.Checked)
			{
				if (UseDelay.Checked)
				{
					LC.openLed(ind,Convert.ToInt32(msDelay.Text));
				}
				else
				{
					LC.openLed(ind);
				}
			}
			else
			{
				LC.closeLed(ind);
			}
			
		}
		
		private void UnCheckAllLeds()
		{
			
		}
		
		private int GetControlID ( object sender)
		{
			string ControlName = ((System.Windows.Forms.Control)sender).Name;
			int ind = Convert.ToInt32(ControlName.Substring(ControlName.Length-1,1));
			return ind;
		}
		
		
		void TrackBarScroll(object sender, EventArgs e)
		{
			int ind = GetControlID(sender);
			TrackBar TB = (TrackBar)sender;
			int Intensity = TB.Value*10;
			LC.changeIntensity(ind,Intensity);
			
		}
		
		
		
	}
}
