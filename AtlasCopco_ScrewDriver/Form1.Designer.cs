namespace AtlasCopco_ScrewDriver
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            this.groupbox_devices = new System.Windows.Forms.GroupBox();
            this.label_usbdevices = new System.Windows.Forms.Label();
            this.combobox_usbdevices = new System.Windows.Forms.ComboBox();
            this.groupbox_save = new System.Windows.Forms.GroupBox();
            this.label_filename = new System.Windows.Forms.Label();
            this.textbox_filename = new System.Windows.Forms.TextBox();
            this.button_browse = new System.Windows.Forms.Button();
            this.label_frequency = new System.Windows.Forms.Label();
            this.textbox_frequency = new System.Windows.Forms.TextBox();
            this.groupbox_status = new System.Windows.Forms.GroupBox();
            this.textbox_status = new System.Windows.Forms.TextBox();
            this.groupbox_connection = new System.Windows.Forms.GroupBox();
            this.button_connect = new System.Windows.Forms.Button();
            this.button_disconnect = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.groupbox_devices.SuspendLayout();
            this.groupbox_save.SuspendLayout();
            this.groupbox_status.SuspendLayout();
            this.groupbox_connection.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupbox_devices
            // 
            this.groupbox_devices.Controls.Add(this.combobox_usbdevices);
            this.groupbox_devices.Controls.Add(this.label_usbdevices);
            this.groupbox_devices.Location = new System.Drawing.Point(13, 11);
            this.groupbox_devices.Name = "groupbox_devices";
            this.groupbox_devices.Size = new System.Drawing.Size(195, 87);
            this.groupbox_devices.TabIndex = 0;
            this.groupbox_devices.TabStop = false;
            this.groupbox_devices.Text = "Devices";
            // 
            // label_usbdevices
            // 
            this.label_usbdevices.AutoSize = true;
            this.label_usbdevices.Location = new System.Drawing.Point(7, 24);
            this.label_usbdevices.Name = "label_usbdevices";
            this.label_usbdevices.Size = new System.Drawing.Size(69, 13);
            this.label_usbdevices.TabIndex = 1;
            this.label_usbdevices.Text = "USB Device:";
            // 
            // combobox_usbdevices
            // 
            this.combobox_usbdevices.FormattingEnabled = true;
            this.combobox_usbdevices.Location = new System.Drawing.Point(82, 19);
            this.combobox_usbdevices.Name = "combobox_usbdevices";
            this.combobox_usbdevices.Size = new System.Drawing.Size(107, 21);
            this.combobox_usbdevices.TabIndex = 1;
            // 
            // groupbox_save
            // 
            this.groupbox_save.Controls.Add(this.textbox_frequency);
            this.groupbox_save.Controls.Add(this.label_frequency);
            this.groupbox_save.Controls.Add(this.button_browse);
            this.groupbox_save.Controls.Add(this.textbox_filename);
            this.groupbox_save.Controls.Add(this.label_filename);
            this.groupbox_save.Location = new System.Drawing.Point(223, 12);
            this.groupbox_save.Name = "groupbox_save";
            this.groupbox_save.Size = new System.Drawing.Size(262, 86);
            this.groupbox_save.TabIndex = 1;
            this.groupbox_save.TabStop = false;
            this.groupbox_save.Text = "Save";
            // 
            // label_filename
            // 
            this.label_filename.AutoSize = true;
            this.label_filename.Location = new System.Drawing.Point(8, 21);
            this.label_filename.Name = "label_filename";
            this.label_filename.Size = new System.Drawing.Size(52, 13);
            this.label_filename.TabIndex = 0;
            this.label_filename.Text = "Filename:";
            // 
            // textbox_filename
            // 
            this.textbox_filename.Location = new System.Drawing.Point(67, 17);
            this.textbox_filename.Multiline = true;
            this.textbox_filename.Name = "textbox_filename";
            this.textbox_filename.Size = new System.Drawing.Size(153, 21);
            this.textbox_filename.TabIndex = 1;
            // 
            // button_browse
            // 
            this.button_browse.Location = new System.Drawing.Point(228, 17);
            this.button_browse.Name = "button_browse";
            this.button_browse.Size = new System.Drawing.Size(24, 20);
            this.button_browse.TabIndex = 2;
            this.button_browse.Text = "...";
            this.button_browse.UseVisualStyleBackColor = true;
            this.button_browse.Click += new System.EventHandler(this.button_browse_Click);
            // 
            // label_frequency
            // 
            this.label_frequency.AutoSize = true;
            this.label_frequency.Location = new System.Drawing.Point(8, 50);
            this.label_frequency.Name = "label_frequency";
            this.label_frequency.Size = new System.Drawing.Size(60, 13);
            this.label_frequency.TabIndex = 3;
            this.label_frequency.Text = "Frequency:";
            // 
            // textbox_frequency
            // 
            this.textbox_frequency.Location = new System.Drawing.Point(67, 48);
            this.textbox_frequency.Multiline = true;
            this.textbox_frequency.Name = "textbox_frequency";
            this.textbox_frequency.Size = new System.Drawing.Size(153, 21);
            this.textbox_frequency.TabIndex = 4;
            // 
            // groupbox_status
            // 
            this.groupbox_status.Controls.Add(this.textbox_status);
            this.groupbox_status.Location = new System.Drawing.Point(13, 115);
            this.groupbox_status.Name = "groupbox_status";
            this.groupbox_status.Size = new System.Drawing.Size(257, 111);
            this.groupbox_status.TabIndex = 2;
            this.groupbox_status.TabStop = false;
            this.groupbox_status.Text = "Status";
            // 
            // textbox_status
            // 
            this.textbox_status.Location = new System.Drawing.Point(9, 16);
            this.textbox_status.Multiline = true;
            this.textbox_status.Name = "textbox_status";
            this.textbox_status.Size = new System.Drawing.Size(241, 86);
            this.textbox_status.TabIndex = 3;
            // 
            // groupbox_connection
            // 
            this.groupbox_connection.Controls.Add(this.button_disconnect);
            this.groupbox_connection.Controls.Add(this.button_connect);
            this.groupbox_connection.Location = new System.Drawing.Point(283, 116);
            this.groupbox_connection.Name = "groupbox_connection";
            this.groupbox_connection.Size = new System.Drawing.Size(201, 109);
            this.groupbox_connection.TabIndex = 3;
            this.groupbox_connection.TabStop = false;
            this.groupbox_connection.Text = "Connection";
            // 
            // button_connect
            // 
            this.button_connect.Location = new System.Drawing.Point(15, 30);
            this.button_connect.Name = "button_connect";
            this.button_connect.Size = new System.Drawing.Size(176, 26);
            this.button_connect.TabIndex = 0;
            this.button_connect.Text = "Connect";
            this.button_connect.UseVisualStyleBackColor = true;
            this.button_connect.Click += new System.EventHandler(this.button_connect_Click);
            // 
            // button_disconnect
            // 
            this.button_disconnect.Location = new System.Drawing.Point(15, 64);
            this.button_disconnect.Name = "button_disconnect";
            this.button_disconnect.Size = new System.Drawing.Size(176, 26);
            this.button_disconnect.TabIndex = 1;
            this.button_disconnect.Text = "Disconnect";
            this.button_disconnect.UseVisualStyleBackColor = true;
            this.button_disconnect.Click += new System.EventHandler(this.button_disconnect_Click);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(497, 234);
            this.Controls.Add(this.groupbox_connection);
            this.Controls.Add(this.groupbox_status);
            this.Controls.Add(this.groupbox_save);
            this.Controls.Add(this.groupbox_devices);
            this.Name = "MainForm";
            this.Text = "AtlasCopco ScrewDriver";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupbox_devices.ResumeLayout(false);
            this.groupbox_devices.PerformLayout();
            this.groupbox_save.ResumeLayout(false);
            this.groupbox_save.PerformLayout();
            this.groupbox_status.ResumeLayout(false);
            this.groupbox_status.PerformLayout();
            this.groupbox_connection.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupbox_devices;
        private System.Windows.Forms.Label label_usbdevices;
        private System.Windows.Forms.ComboBox combobox_usbdevices;
        private System.Windows.Forms.GroupBox groupbox_save;
        private System.Windows.Forms.Label label_filename;
        private System.Windows.Forms.TextBox textbox_filename;
        private System.Windows.Forms.Button button_browse;
        private System.Windows.Forms.TextBox textbox_frequency;
        private System.Windows.Forms.Label label_frequency;
        private System.Windows.Forms.GroupBox groupbox_status;
        private System.Windows.Forms.TextBox textbox_status;
        private System.Windows.Forms.GroupBox groupbox_connection;
        private System.Windows.Forms.Button button_disconnect;
        private System.Windows.Forms.Button button_connect;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Timer timer1;
    }
}

