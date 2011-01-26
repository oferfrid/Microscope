/*
 * Created by SharpDevelop.
 * User: oferfrid
 * Date: 22/03/2009
 * Time: 13:16
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace LedTest
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
			this.trackBar1 = new System.Windows.Forms.TrackBar();
			this.trackBar2 = new System.Windows.Forms.TrackBar();
			this.trackBar3 = new System.Windows.Forms.TrackBar();
			this.trackBar4 = new System.Windows.Forms.TrackBar();
			this.ConnectDisconnect = new System.Windows.Forms.CheckBox();
			this.UseDelay = new System.Windows.Forms.CheckBox();
			this.label1 = new System.Windows.Forms.Label();
			this.msDelay = new System.Windows.Forms.TextBox();
			this.Led1 = new System.Windows.Forms.CheckBox();
			this.Led2 = new System.Windows.Forms.CheckBox();
			this.Led3 = new System.Windows.Forms.CheckBox();
			this.Led4 = new System.Windows.Forms.CheckBox();
			((System.ComponentModel.ISupportInitialize)(this.trackBar1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.trackBar2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.trackBar3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.trackBar4)).BeginInit();
			this.SuspendLayout();
			// 
			// trackBar1
			// 
			this.trackBar1.Location = new System.Drawing.Point(65, 68);
			this.trackBar1.Name = "trackBar1";
			this.trackBar1.Size = new System.Drawing.Size(104, 45);
			this.trackBar1.TabIndex = 5;
			this.trackBar1.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// trackBar2
			// 
			this.trackBar2.Location = new System.Drawing.Point(65, 105);
			this.trackBar2.Name = "trackBar2";
			this.trackBar2.Size = new System.Drawing.Size(104, 45);
			this.trackBar2.TabIndex = 5;
			this.trackBar2.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// trackBar3
			// 
			this.trackBar3.Location = new System.Drawing.Point(65, 142);
			this.trackBar3.Name = "trackBar3";
			this.trackBar3.Size = new System.Drawing.Size(104, 45);
			this.trackBar3.TabIndex = 5;
			this.trackBar3.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// trackBar4
			// 
			this.trackBar4.Location = new System.Drawing.Point(65, 179);
			this.trackBar4.Name = "trackBar4";
			this.trackBar4.Size = new System.Drawing.Size(104, 45);
			this.trackBar4.TabIndex = 5;
			this.trackBar4.Scroll += new System.EventHandler(this.TrackBarScroll);
			// 
			// ConnectDisconnect
			// 
			this.ConnectDisconnect.Appearance = System.Windows.Forms.Appearance.Button;
			this.ConnectDisconnect.Location = new System.Drawing.Point(12, 12);
			this.ConnectDisconnect.Name = "ConnectDisconnect";
			this.ConnectDisconnect.Size = new System.Drawing.Size(157, 24);
			this.ConnectDisconnect.TabIndex = 6;
			this.ConnectDisconnect.Text = "Connect";
			this.ConnectDisconnect.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.ConnectDisconnect.UseVisualStyleBackColor = true;
			this.ConnectDisconnect.CheckedChanged += new System.EventHandler(this.ConnectDisconnectClick);
			// 
			// UseDelay
			// 
			this.UseDelay.Location = new System.Drawing.Point(13, 38);
			this.UseDelay.Name = "UseDelay";
			this.UseDelay.Size = new System.Drawing.Size(104, 24);
			this.UseDelay.TabIndex = 7;
			this.UseDelay.Text = "Use Delay";
			this.UseDelay.UseVisualStyleBackColor = true;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(140, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 23);
			this.label1.TabIndex = 8;
			this.label1.Text = "(ms)";
			// 
			// msDelay
			// 
			this.msDelay.Location = new System.Drawing.Point(85, 40);
			this.msDelay.Name = "msDelay";
			this.msDelay.Size = new System.Drawing.Size(49, 20);
			this.msDelay.TabIndex = 9;
			this.msDelay.Text = "15";
			// 
			// Led1
			// 
			this.Led1.Location = new System.Drawing.Point(13, 68);
			this.Led1.Name = "Led1";
			this.Led1.Size = new System.Drawing.Size(55, 24);
			this.Led1.TabIndex = 10;
			this.Led1.Text = "Led1";
			this.Led1.UseVisualStyleBackColor = true;
			this.Led1.CheckedChanged += new System.EventHandler(this.LedCheckedChanged);
			// 
			// Led2
			// 
			this.Led2.Location = new System.Drawing.Point(13, 105);
			this.Led2.Name = "Led2";
			this.Led2.Size = new System.Drawing.Size(55, 24);
			this.Led2.TabIndex = 10;
			this.Led2.Text = "Led2";
			this.Led2.UseVisualStyleBackColor = true;
			this.Led2.CheckedChanged += new System.EventHandler(this.LedCheckedChanged);
			// 
			// Led3
			// 
			this.Led3.Location = new System.Drawing.Point(13, 142);
			this.Led3.Name = "Led3";
			this.Led3.Size = new System.Drawing.Size(55, 24);
			this.Led3.TabIndex = 10;
			this.Led3.Text = "Led2";
			this.Led3.UseVisualStyleBackColor = true;
			this.Led3.CheckedChanged += new System.EventHandler(this.LedCheckedChanged);
			// 
			// Led4
			// 
			this.Led4.Location = new System.Drawing.Point(13, 179);
			this.Led4.Name = "Led4";
			this.Led4.Size = new System.Drawing.Size(55, 24);
			this.Led4.TabIndex = 10;
			this.Led4.Text = "Led3";
			this.Led4.UseVisualStyleBackColor = true;
			this.Led4.CheckedChanged += new System.EventHandler(this.LedCheckedChanged);
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(178, 241);
			this.Controls.Add(this.Led4);
			this.Controls.Add(this.Led3);
			this.Controls.Add(this.Led2);
			this.Controls.Add(this.Led1);
			this.Controls.Add(this.msDelay);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.UseDelay);
			this.Controls.Add(this.ConnectDisconnect);
			this.Controls.Add(this.trackBar4);
			this.Controls.Add(this.trackBar3);
			this.Controls.Add(this.trackBar2);
			this.Controls.Add(this.trackBar1);
			this.Name = "MainForm";
			this.Text = "LedTest";
			this.Load += new System.EventHandler(this.MainFormLoad);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainFormFormClosing);
			((System.ComponentModel.ISupportInitialize)(this.trackBar1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.trackBar2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.trackBar3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.trackBar4)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();
		}
		private System.Windows.Forms.CheckBox Led2;
		private System.Windows.Forms.CheckBox Led3;
		private System.Windows.Forms.CheckBox Led4;
		private System.Windows.Forms.CheckBox Led1;
		private System.Windows.Forms.TextBox msDelay;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.CheckBox UseDelay;
		private System.Windows.Forms.TrackBar trackBar4;
		private System.Windows.Forms.TrackBar trackBar3;
		private System.Windows.Forms.TrackBar trackBar2;
		private System.Windows.Forms.TrackBar trackBar1;
		private System.Windows.Forms.CheckBox ConnectDisconnect;
	}
}
