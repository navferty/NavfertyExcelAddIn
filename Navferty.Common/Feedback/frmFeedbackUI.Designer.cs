﻿namespace Navferty.Common.Feedback
{
	partial class frmFeedbackUI
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.btnSend = new System.Windows.Forms.Button();
            this.lblMessage = new System.Windows.Forms.Label();
            this.chkIncludeScreenshots = new System.Windows.Forms.CheckBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.txtUserMessage = new System.Windows.Forms.TextBox();
            this.lblSummary = new System.Windows.Forms.LinkLabel();
            this.llGotoGithub = new System.Windows.Forms.LinkLabel();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSend
            // 
            this.btnSend.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnSend.Location = new System.Drawing.Point(428, 231);
            this.btnSend.Name = "btnSend";
            this.tableLayoutPanel1.SetRowSpan(this.btnSend, 2);
            this.btnSend.Size = new System.Drawing.Size(113, 34);
            this.btnSend.TabIndex = 1;
            this.btnSend.Text = "Send";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.OnSend);
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.lblMessage, 2);
            this.lblMessage.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblMessage.Location = new System.Drawing.Point(7, 4);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(534, 13);
            this.lblMessage.TabIndex = 1;
            this.lblMessage.Text = "message";
            // 
            // chkIncludeScreenshots
            // 
            this.chkIncludeScreenshots.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.chkIncludeScreenshots, 2);
            this.chkIncludeScreenshots.Dock = System.Windows.Forms.DockStyle.Top;
            this.chkIncludeScreenshots.Location = new System.Drawing.Point(7, 208);
            this.chkIncludeScreenshots.Name = "chkIncludeScreenshots";
            this.chkIncludeScreenshots.Size = new System.Drawing.Size(534, 17);
            this.chkIncludeScreenshots.TabIndex = 2;
            this.chkIncludeScreenshots.Text = "include screenshot";
            this.chkIncludeScreenshots.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.btnSend, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.lblMessage, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.txtUserMessage, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.chkIncludeScreenshots, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.lblSummary, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.llGotoGithub, 0, 4);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(8, 8);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.Padding = new System.Windows.Forms.Padding(4);
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.Size = new System.Drawing.Size(548, 272);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // txtUserMessage
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.txtUserMessage, 2);
            this.txtUserMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtUserMessage.HideSelection = false;
            this.txtUserMessage.Location = new System.Drawing.Point(7, 20);
            this.txtUserMessage.MaxLength = 1000;
            this.txtUserMessage.Multiline = true;
            this.txtUserMessage.Name = "txtUserMessage";
            this.txtUserMessage.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtUserMessage.Size = new System.Drawing.Size(534, 182);
            this.txtUserMessage.TabIndex = 0;
            // 
            // lblSummary
            // 
            this.lblSummary.AutoSize = true;
            this.lblSummary.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblSummary.Location = new System.Drawing.Point(7, 228);
            this.lblSummary.Name = "lblSummary";
            this.lblSummary.Size = new System.Drawing.Size(415, 13);
            this.lblSummary.TabIndex = 4;
            this.lblSummary.TabStop = true;
            this.lblSummary.Text = "summary";
            // 
            // llGotoGithub
            // 
            this.llGotoGithub.AutoSize = true;
            this.llGotoGithub.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.llGotoGithub.Location = new System.Drawing.Point(7, 255);
            this.llGotoGithub.Name = "llGotoGithub";
            this.llGotoGithub.Size = new System.Drawing.Size(415, 13);
            this.llGotoGithub.TabIndex = 2;
            this.llGotoGithub.TabStop = true;
            this.llGotoGithub.Text = "goto github";
            this.llGotoGithub.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.OnGotoGithub);
            // 
            // frmFeedbackUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(564, 288);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmFeedbackUI";
            this.Padding = new System.Windows.Forms.Padding(8);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "frmFeedbackUI";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnSend;
		private System.Windows.Forms.Label lblMessage;
		private System.Windows.Forms.CheckBox chkIncludeScreenshots;
		private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
		private System.Windows.Forms.TextBox txtUserMessage;
		private System.Windows.Forms.LinkLabel lblSummary;
		private System.Windows.Forms.LinkLabel llGotoGithub;
	}
}