using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Reflection;

namespace WirecastCSharp
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.ListBox layerShotList;
		private System.Windows.Forms.Button button1;
        private CheckBox autoLiveCheckbox;
        private Button broadcastButton;
        private Button recordButton;
        private ContextMenuStrip contextMenuStrip1;
        private ToolStripMenuItem masterLayer1ToolStripMenuItem;
        private ToolStripMenuItem masterLayer2ToolStripMenuItem;
        private ToolStripMenuItem masterLayer3ToolStripMenuItem;
        private ToolStripMenuItem masterLayer4ToolStripMenuItem;
        private ToolStripMenuItem masterLayer5ToolStripMenuItem;
        private IContainer components;
        private ComboBox activeTransitionBox;
        private CheckBox audioMute;
        private Button snapShotButton;
        private SaveFileDialog saveSnapShotDialog;
        private Button removeShot;
        private ToolStripSeparator toolStripMenuItem1;
        private ToolStripMenuItem addMediaToolStripMenuItem;
        private ToolStripMenuItem removeMediaToolStripMenuItem;
        private OpenFileDialog selectMediaDialog;
        private ToolStripMenuItem visibleToolStripMenuItem;
        private TextBox shotNameTextField;
        private ComboBox transitionSpeedBox;
        private Label layerNameText;
        private Button refreshLayer;
        private Label previewShotLabel;
        private Label liveShotLabel;

        private BindingList<ShotItem> shotList;

        //
        private WirecastBinding _Wirecast;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

            _Wirecast = new WirecastBinding();
            _Wirecast.Initialize();

            shotList = new BindingList<ShotItem>();
            layerShotList.DataSource = shotList;
            layerShotList.DisplayMember = "statusName";
		}

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

            _Wirecast = null;
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.layerShotList = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.masterLayer1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.masterLayer2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.masterLayer3ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.masterLayer4ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.masterLayer5ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.addMediaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.removeMediaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.visibleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button1 = new System.Windows.Forms.Button();
            this.autoLiveCheckbox = new System.Windows.Forms.CheckBox();
            this.broadcastButton = new System.Windows.Forms.Button();
            this.recordButton = new System.Windows.Forms.Button();
            this.activeTransitionBox = new System.Windows.Forms.ComboBox();
            this.audioMute = new System.Windows.Forms.CheckBox();
            this.snapShotButton = new System.Windows.Forms.Button();
            this.saveSnapShotDialog = new System.Windows.Forms.SaveFileDialog();
            this.removeShot = new System.Windows.Forms.Button();
            this.selectMediaDialog = new System.Windows.Forms.OpenFileDialog();
            this.shotNameTextField = new System.Windows.Forms.TextBox();
            this.transitionSpeedBox = new System.Windows.Forms.ComboBox();
            this.layerNameText = new System.Windows.Forms.Label();
            this.refreshLayer = new System.Windows.Forms.Button();
            this.previewShotLabel = new System.Windows.Forms.Label();
            this.liveShotLabel = new System.Windows.Forms.Label();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // layerShotList
            // 
            this.layerShotList.ContextMenuStrip = this.contextMenuStrip1;
            this.layerShotList.ItemHeight = 20;
            this.layerShotList.Location = new System.Drawing.Point(11, 53);
            this.layerShotList.Name = "layerShotList";
            this.layerShotList.Size = new System.Drawing.Size(1453, 464);
            this.layerShotList.TabIndex = 0;
            this.layerShotList.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.masterLayer1ToolStripMenuItem,
            this.masterLayer2ToolStripMenuItem,
            this.masterLayer3ToolStripMenuItem,
            this.masterLayer4ToolStripMenuItem,
            this.masterLayer5ToolStripMenuItem,
            this.toolStripMenuItem1,
            this.addMediaToolStripMenuItem,
            this.removeMediaToolStripMenuItem,
            this.visibleToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(216, 250);
            // 
            // masterLayer1ToolStripMenuItem
            // 
            this.masterLayer1ToolStripMenuItem.Name = "masterLayer1ToolStripMenuItem";
            this.masterLayer1ToolStripMenuItem.Size = new System.Drawing.Size(215, 30);
            this.masterLayer1ToolStripMenuItem.Text = "Master Layer 1";
            // 
            // masterLayer2ToolStripMenuItem
            // 
            this.masterLayer2ToolStripMenuItem.Name = "masterLayer2ToolStripMenuItem";
            this.masterLayer2ToolStripMenuItem.Size = new System.Drawing.Size(215, 30);
            this.masterLayer2ToolStripMenuItem.Text = "Master Layer 2";
            // 
            // masterLayer3ToolStripMenuItem
            // 
            this.masterLayer3ToolStripMenuItem.Name = "masterLayer3ToolStripMenuItem";
            this.masterLayer3ToolStripMenuItem.Size = new System.Drawing.Size(215, 30);
            this.masterLayer3ToolStripMenuItem.Text = "Master Layer 3";
            // 
            // masterLayer4ToolStripMenuItem
            // 
            this.masterLayer4ToolStripMenuItem.Name = "masterLayer4ToolStripMenuItem";
            this.masterLayer4ToolStripMenuItem.Size = new System.Drawing.Size(215, 30);
            this.masterLayer4ToolStripMenuItem.Text = "Master Layer 4";
            // 
            // masterLayer5ToolStripMenuItem
            // 
            this.masterLayer5ToolStripMenuItem.Name = "masterLayer5ToolStripMenuItem";
            this.masterLayer5ToolStripMenuItem.Size = new System.Drawing.Size(215, 30);
            this.masterLayer5ToolStripMenuItem.Text = "Master Layer 5";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(212, 6);
            // 
            // addMediaToolStripMenuItem
            // 
            this.addMediaToolStripMenuItem.Name = "addMediaToolStripMenuItem";
            this.addMediaToolStripMenuItem.Size = new System.Drawing.Size(215, 30);
            this.addMediaToolStripMenuItem.Text = "Add Media...";
            this.addMediaToolStripMenuItem.Click += new System.EventHandler(this.addMediaToolStripMenuItem_Click);
            // 
            // removeMediaToolStripMenuItem
            // 
            this.removeMediaToolStripMenuItem.Name = "removeMediaToolStripMenuItem";
            this.removeMediaToolStripMenuItem.Size = new System.Drawing.Size(215, 30);
            this.removeMediaToolStripMenuItem.Text = "Remove Media...";
            this.removeMediaToolStripMenuItem.Click += new System.EventHandler(this.removeMediaToolStripMenuItem_Click);
            // 
            // visibleToolStripMenuItem
            // 
            this.visibleToolStripMenuItem.Name = "visibleToolStripMenuItem";
            this.visibleToolStripMenuItem.Size = new System.Drawing.Size(215, 30);
            this.visibleToolStripMenuItem.Text = "Visible";
            this.visibleToolStripMenuItem.Click += new System.EventHandler(this.visibleToolStripMenuItem_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(822, 537);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 30);
            this.button1.TabIndex = 1;
            this.button1.Text = "Go";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // autoLiveCheckbox
            // 
            this.autoLiveCheckbox.AutoSize = true;
            this.autoLiveCheckbox.Location = new System.Drawing.Point(930, 540);
            this.autoLiveCheckbox.Name = "autoLiveCheckbox";
            this.autoLiveCheckbox.Size = new System.Drawing.Size(97, 24);
            this.autoLiveCheckbox.TabIndex = 2;
            this.autoLiveCheckbox.Text = "AutoLive";
            this.autoLiveCheckbox.UseVisualStyleBackColor = true;
            this.autoLiveCheckbox.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // broadcastButton
            // 
            this.broadcastButton.Location = new System.Drawing.Point(12, 18);
            this.broadcastButton.Name = "broadcastButton";
            this.broadcastButton.Size = new System.Drawing.Size(168, 29);
            this.broadcastButton.TabIndex = 3;
            this.broadcastButton.Text = "broadcast";
            this.broadcastButton.UseVisualStyleBackColor = true;
            this.broadcastButton.Click += new System.EventHandler(this.button2_Click);
            // 
            // recordButton
            // 
            this.recordButton.Location = new System.Drawing.Point(185, 18);
            this.recordButton.Name = "recordButton";
            this.recordButton.Size = new System.Drawing.Size(167, 29);
            this.recordButton.TabIndex = 4;
            this.recordButton.Text = "record";
            this.recordButton.UseVisualStyleBackColor = true;
            this.recordButton.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // activeTransitionBox
            // 
            this.activeTransitionBox.BackColor = System.Drawing.SystemColors.Window;
            this.activeTransitionBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.activeTransitionBox.FormattingEnabled = true;
            this.activeTransitionBox.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.activeTransitionBox.Items.AddRange(new object[] {
            "Left Popup",
            "Right Popup"});
            this.activeTransitionBox.Location = new System.Drawing.Point(512, 538);
            this.activeTransitionBox.MaxDropDownItems = 2;
            this.activeTransitionBox.Name = "activeTransitionBox";
            this.activeTransitionBox.Size = new System.Drawing.Size(145, 28);
            this.activeTransitionBox.TabIndex = 8;
            this.activeTransitionBox.SelectedIndexChanged += new System.EventHandler(this.activeTransitionBox_SelectedIndexChanged);
            // 
            // audioMute
            // 
            this.audioMute.AutoSize = true;
            this.audioMute.Location = new System.Drawing.Point(1321, 540);
            this.audioMute.Name = "audioMute";
            this.audioMute.Size = new System.Drawing.Size(143, 24);
            this.audioMute.TabIndex = 9;
            this.audioMute.Text = "Mute Speakers";
            this.audioMute.UseVisualStyleBackColor = true;
            this.audioMute.CheckedChanged += new System.EventHandler(this.audioMute_CheckedChanged);
            // 
            // snapShotButton
            // 
            this.snapShotButton.Location = new System.Drawing.Point(1321, 18);
            this.snapShotButton.Name = "snapShotButton";
            this.snapShotButton.Size = new System.Drawing.Size(143, 29);
            this.snapShotButton.TabIndex = 10;
            this.snapShotButton.Text = "Snapshot...";
            this.snapShotButton.UseVisualStyleBackColor = true;
            this.snapShotButton.Click += new System.EventHandler(this.snapShotButton_Click);
            // 
            // saveSnapShotDialog
            // 
            this.saveSnapShotDialog.DefaultExt = "jpg";
            this.saveSnapShotDialog.Filter = "JPEG files (*.jpg)|*.jpg";
            // 
            // removeShot
            // 
            this.removeShot.Location = new System.Drawing.Point(11, 537);
            this.removeShot.Name = "removeShot";
            this.removeShot.Size = new System.Drawing.Size(128, 29);
            this.removeShot.TabIndex = 11;
            this.removeShot.Text = "Remove Shot";
            this.removeShot.UseVisualStyleBackColor = true;
            this.removeShot.Click += new System.EventHandler(this.button3_Click);
            // 
            // shotNameTextField
            // 
            this.shotNameTextField.Location = new System.Drawing.Point(146, 538);
            this.shotNameTextField.Name = "shotNameTextField";
            this.shotNameTextField.Size = new System.Drawing.Size(272, 26);
            this.shotNameTextField.TabIndex = 12;
            this.shotNameTextField.Text = "Selected Shot Name";
            this.shotNameTextField.TextChanged += new System.EventHandler(this.shotNameTextField_TextChanged);
            // 
            // transitionSpeedBox
            // 
            this.transitionSpeedBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.transitionSpeedBox.FormattingEnabled = true;
            this.transitionSpeedBox.Items.AddRange(new object[] {
            "slowest",
            "slow",
            "normal",
            "faster",
            "fastest"});
            this.transitionSpeedBox.Location = new System.Drawing.Point(664, 538);
            this.transitionSpeedBox.Name = "transitionSpeedBox";
            this.transitionSpeedBox.Size = new System.Drawing.Size(152, 28);
            this.transitionSpeedBox.TabIndex = 13;
            this.transitionSpeedBox.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // layerNameText
            // 
            this.layerNameText.AutoSize = true;
            this.layerNameText.Location = new System.Drawing.Point(660, 22);
            this.layerNameText.Name = "layerNameText";
            this.layerNameText.Size = new System.Drawing.Size(161, 20);
            this.layerNameText.TabIndex = 14;
            this.layerNameText.Text = "Selected Layer Name";
            this.layerNameText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // refreshLayer
            // 
            this.refreshLayer.Location = new System.Drawing.Point(1166, 18);
            this.refreshLayer.Name = "refreshLayer";
            this.refreshLayer.Size = new System.Drawing.Size(149, 29);
            this.refreshLayer.TabIndex = 15;
            this.refreshLayer.Text = "Refresh Status";
            this.refreshLayer.UseVisualStyleBackColor = true;
            this.refreshLayer.Click += new System.EventHandler(this.refreshLayer_Click);
            // 
            // previewShotLabel
            // 
            this.previewShotLabel.AutoSize = true;
            this.previewShotLabel.Location = new System.Drawing.Point(417, 22);
            this.previewShotLabel.Name = "previewShotLabel";
            this.previewShotLabel.Size = new System.Drawing.Size(105, 20);
            this.previewShotLabel.TabIndex = 16;
            this.previewShotLabel.Text = "Preview Shot:";
            // 
            // liveShotLabel
            // 
            this.liveShotLabel.AutoSize = true;
            this.liveShotLabel.Location = new System.Drawing.Point(913, 22);
            this.liveShotLabel.Name = "liveShotLabel";
            this.liveShotLabel.Size = new System.Drawing.Size(79, 20);
            this.liveShotLabel.TabIndex = 17;
            this.liveShotLabel.Text = "Live Shot:";
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(8, 19);
            this.CausesValidation = false;
            this.ClientSize = new System.Drawing.Size(1477, 579);
            this.Controls.Add(this.liveShotLabel);
            this.Controls.Add(this.previewShotLabel);
            this.Controls.Add(this.refreshLayer);
            this.Controls.Add(this.layerNameText);
            this.Controls.Add(this.transitionSpeedBox);
            this.Controls.Add(this.shotNameTextField);
            this.Controls.Add(this.removeShot);
            this.Controls.Add(this.snapShotButton);
            this.Controls.Add(this.audioMute);
            this.Controls.Add(this.activeTransitionBox);
            this.Controls.Add(this.recordButton);
            this.Controls.Add(this.broadcastButton);
            this.Controls.Add(this.autoLiveCheckbox);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.layerShotList);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

        private void LoadLayerShots()
        {
            shotList.Clear();

            //  Populate Listbox with shots
            int count = _Wirecast.GetShotCount();
            for (int idx = 1; idx <= count; idx++)
            {
                string name = _Wirecast.GetShotNameWithIndex(idx);
                bool isInPreview = _Wirecast.IsShotNameInPreview(name);
                bool isLive = _Wirecast.IsShotNameLive(name);
                ShotItem sItem = new ShotItem();
                sItem.name = name;
                sItem.id = _Wirecast.GetShotIDByName(name);
                sItem.UpdateStatusName(isLive, isInPreview);
                shotList.Add(sItem);

                if (_Wirecast.IsActiveShot(name))
                {
                    layerShotList.SelectedIndex = idx - 1;
                }
            }

            visibleToolStripMenuItem.Checked = _Wirecast.IsLayerVisible();

            int layerIndex = _Wirecast.GetSelectedLayerIndex();
            masterLayer1ToolStripMenuItem.Checked = layerIndex == 1;
            masterLayer2ToolStripMenuItem.Checked = layerIndex == 2;
            masterLayer3ToolStripMenuItem.Checked = layerIndex == 3;
            masterLayer4ToolStripMenuItem.Checked = layerIndex == 4;
            masterLayer5ToolStripMenuItem.Checked = layerIndex == 5;

            //  Update Selected Layer Name Title
            layerNameText.Text = "Selected Layer: " + _Wirecast.GetSelectedLayerName().ToUpper();
        }

		private void Form1_Load(object sender, System.EventArgs e)
		{
            //  Initialize the Context Menu Events
            masterLayer1ToolStripMenuItem.Click += new System.EventHandler(this.masterLayer1_Click);
            masterLayer2ToolStripMenuItem.Click += new System.EventHandler(this.masterLayer2_Click);
            masterLayer3ToolStripMenuItem.Click += new System.EventHandler(this.masterLayer3_Click);
            masterLayer4ToolStripMenuItem.Click += new System.EventHandler(this.masterLayer4_Click);
            masterLayer5ToolStripMenuItem.Click += new System.EventHandler(this.masterLayer5_Click);

			try {
                LoadLayerShots();

                //  Read Broadcasting state
                if (_Wirecast.IsBroadcasting())
                {
                    broadcastButton.Text = "Stop Broadcasting";
                }
                else
                {
                    broadcastButton.Text = "Start Broadcasting";
                }

                //  Read Recording State
                if (_Wirecast.IsRecording())
                {
                    recordButton.Text = "Stop Recording";
                }
                else
                {
                    recordButton.Text = "Start Recording";
                }

                //  Read Transition Speed
                transitionSpeedBox.SelectedItem = _Wirecast.GetTransitionSpeed();

                //  Read Active Transition
                activeTransitionBox.SelectedIndex = _Wirecast.GetActiveTransitionIndex();

                //  Read Audio Mute state
                audioMute.Checked = _Wirecast.IsAudioMutedToSpeakers();

                //  Audo Live
                autoLiveCheckbox.Checked = _Wirecast.IsAutoLiveOn();
			} catch( Exception theException ) {
				String errorMessage;
				errorMessage = "Error: ";
				errorMessage = String.Concat( errorMessage, theException.Message );
				errorMessage = String.Concat( errorMessage, " Line: " );
				errorMessage = String.Concat( errorMessage, theException.Source );

				MessageBox.Show( errorMessage, "Error" );
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
            _Wirecast.Go();

            refreshLivePreviewShots();
		}

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (autoLiveCheckbox.Checked != _Wirecast.IsAutoLiveOn())
            {
                _Wirecast.ToggleAutoLive();
                autoLiveCheckbox.Checked = _Wirecast.IsAutoLiveOn();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _Wirecast.ToggleBroadcast();
            if (_Wirecast.IsBroadcasting())
            {
                broadcastButton.Text = "Stop Broadcasting";
            }
            else
            {
                broadcastButton.Text = "Start Broadcasting";
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            _Wirecast.ToggleRecord();
            if (_Wirecast.IsRecording())
            {
                recordButton.Text = "Stop Recording";
            }
            else
            {
                recordButton.Text = "Start Recording";
            }
        }

        private void masterLayer1_Click(object sender, System.EventArgs e)
        {
            _Wirecast.SwitchLayer(1);
            LoadLayerShots();
        }

        private void masterLayer2_Click(object sender, System.EventArgs e)
        {
            _Wirecast.SwitchLayer(2);
            LoadLayerShots();
        }

        private void masterLayer3_Click(object sender, System.EventArgs e)
        {
            _Wirecast.SwitchLayer(3);
            LoadLayerShots();
        }

        private void masterLayer4_Click(object sender, System.EventArgs e)
        {
            _Wirecast.SwitchLayer(4);
            LoadLayerShots();
        }

        private void masterLayer5_Click(object sender, System.EventArgs e)
        {
            _Wirecast.SwitchLayer(5);
            LoadLayerShots();
        }

        private void setTransitionSpeed(string speed)
        {
            _Wirecast.SetTransitionSpeed(speed);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            setTransitionSpeed((string)transitionSpeedBox.SelectedItem);
        }

        private void activeTransitionBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selIndex = activeTransitionBox.SelectedIndex + 1;
            _Wirecast.SetActiveTransitionIndex(selIndex);
        }

        private void audioMute_CheckedChanged(object sender, EventArgs e)
        {
            if (audioMute.Checked != _Wirecast.IsAudioMutedToSpeakers())
            {
                _Wirecast.ToggleAudioMutedToSpeakers();
            }
        }

        private void snapShotButton_Click(object sender, EventArgs e)
        {
            if (saveSnapShotDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveSnapShotDialog.FileName;
                if (filePath != "")
                {
                    _Wirecast.SaveSnapshot(filePath);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ShotItem sItem = (ShotItem)layerShotList.SelectedItem;
            _Wirecast.RemoveShotWithName(sItem.name);
            LoadLayerShots();
        }

        private void addMediaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (selectMediaDialog.ShowDialog() == DialogResult.OK)
            {
                string mediaPath = selectMediaDialog.FileName;
                if (mediaPath != "")
                {
                    _Wirecast.AddShotWithMedia(mediaPath);
                    LoadLayerShots();
                }
            }
        }

        private void removeMediaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (selectMediaDialog.ShowDialog() == DialogResult.OK)
            {
                string mediaPath = selectMediaDialog.FileName;

                if(mediaPath != "")
                {
                    _Wirecast.RemoveMedia(mediaPath);
                    LoadLayerShots();
                }
            }
        }

        private void visibleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _Wirecast.ToggleLayerVisibility();
            visibleToolStripMenuItem.Checked = _Wirecast.IsLayerVisible();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShotItem sItem = (ShotItem)layerShotList.SelectedItem;
            shotNameTextField.Text = sItem.name;
            _Wirecast.SetActiveShot(sItem.name);

            bool isInPreview = _Wirecast.IsShotNameInPreview(sItem.name);
            bool isLive = _Wirecast.IsShotNameLive(sItem.name);
            sItem.UpdateStatusName(isLive, isInPreview);

            refreshLivePreviewShots();
        }

        private void shotNameTextField_TextChanged(object sender, EventArgs e)
        {
            ShotItem sItem = (ShotItem)layerShotList.SelectedItem;
            if (sItem.name != shotNameTextField.Text)
            {
                int shotID = _Wirecast.GetShotIDByName(sItem.name);
                _Wirecast.SetShotNameWithName(sItem.name, shotNameTextField.Text);
                LoadLayerShots();
            }
        }

        private void refreshLayer_Click(object sender, EventArgs e)
        {
            LoadLayerShots();

            //  Update Current Live and Preview Shots
            refreshLivePreviewShots();
        }

        private void refreshLivePreviewShots()
        {
            string previewShotName = _Wirecast.GetPreviewShotName();
            string liveShotName = _Wirecast.GetLiveShotName();

            previewShotLabel.Text = previewShotName;
            liveShotLabel.Text = liveShotName;
        }
	}

    public class ShotItem
    {
        public string name { get; set; }
        public int id { get; set; }
        public string statusName { get; set; }
        public void UpdateStatusName(bool live, bool preview)
        {
            string previewStr = preview ? "YES" : "NO";
            string liveStr = live ? "YES" : "NO";
            statusName = name + "\t\t(PREVIEW: " + previewStr + ", LIVE: " + liveStr + ")";
        }
    }


    /// <summary>
    /// Proxy Object for the Late binding
    /// Provide more logical API for bindings
    /// </summary>
    public class WirecastBinding
    {
        private object _Wirecast;
        private object _Document;
        private object _SelectedLayer;
        private int _SelectedLayerIndex;

        /// <summary>
        /// Late binding helper class.
        /// static bindings to help you get/set via OLE/COM.
        /// </summary>
        class Late
        {
            public static void Set(object obj, string sProperty, object oValue)
            {
                object[] oParam = new object[1];
                oParam[0] = oValue;
                obj.GetType().InvokeMember(sProperty, BindingFlags.SetProperty, null, obj, oParam);
            }
            public static object Get(object obj, string sProperty, object oValue)
            {
                object[] oParam = new object[1];
                oParam[0] = oValue;
                return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, oParam);
            }
            public static object Get(object obj, string sProperty, object[] oValue)
            {
                return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, oValue);
            }
            public static object Get(object obj, string sProperty)
            {
                return obj.GetType().InvokeMember(sProperty, BindingFlags.GetProperty, null, obj, null);
            }
            public static object Invoke(object obj, string sProperty, object[] oParam)
            {
                return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
            }
            public static object Invoke(object obj, string sProperty, object oValue)
            {
                object[] oParam = new object[1];
                oParam[0] = oValue;
                return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
            }
            public static object Invoke(object obj, string sProperty, object oValue, object oValue2)
            {
                object[] oParam = new object[2];
                oParam[0] = oValue;
                oParam[1] = oValue2;
                return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, oParam);
            }
            public static object Invoke(object obj, string sProperty)
            {
                return obj.GetType().InvokeMember(sProperty, BindingFlags.InvokeMethod, null, obj, null);
            }
        }

        /// <summary>
        /// Initialize the Binding object with Document Index 1
        /// And the Normal Layer (3) Selected
        /// </summary>
        public void Initialize()
        {
            _Wirecast = null;
            _Document = null;
            _SelectedLayer = null;
            _SelectedLayerIndex = -1;
            try
            {
                _Wirecast = Marshal.GetActiveObject("Wirecast.Application");
            }
            catch
            {
                Type objClassType = Type.GetTypeFromProgID("Wirecast.Application");
                _Wirecast = Activator.CreateInstance(objClassType);
            }

            SwitchDocument(1);
        }

        /// <summary>
        /// Returns a string array of the layer names
        /// </summary>  
        public string[] ValidLayerNames()
        {
            return new string[] { "text", "overlay", "normal", "underlay", "audio" };
        }

        /// <summary>
        /// Converts a given Layer Name to the associated Index value
        /// If the name is invalid, returns -1
        /// </summary>
        public int LayerNameToIndex(string layer)
        {
            string[] layerNames = ValidLayerNames();
            for (int i = 0; i < 5; i++)
            {
                if (layerNames[i] == layer)
                {
                    return i + 1;
                }
            }

            return -1;
        }
        /// <summary>
        /// Get the Name of the layer associated with the index value.
        /// If the index is out of bounds, returns an empty string
        /// </summary>
        public string LayerIndexToName(int index)
        {
            if (index >= 1 && index <= 5)
            {
                string[] layerNames = ValidLayerNames();
                return layerNames[index - 1];
            }

            return "";
        }

        /// <summary>
        /// Switches the internal document of the binding object to the new one
        /// The seletec layer of the new document will be the same as the previous document
        /// So if the selected layer was "text"/1, then that is the selected layer of the new document
        /// All future call to the Document and Layer API will apply to the new document
        /// </summary>
        public bool SwitchDocument(int index)
        {
            bool validDoc = false;
            object newDocument = Late.Invoke(_Wirecast, "DocumentByIndex", index);

            if (newDocument != null)
            {
                _Document = newDocument;
                validDoc = true;

                //  After Change the Document, we need to update the layer
                SwitchLayer(3);
            }

            return validDoc;
        }

        /// <summary>
        /// Same as SwitchDocument with an index except indead of a document index, use the document name
        /// </summary>
        public bool SwitchDocument(string docName)
        {
            bool validDoc = false;
            object newDocument = Late.Invoke(_Wirecast, "DocumentByName", docName);

            if (newDocument != null)
            {
                _Document = newDocument;
                validDoc = true;

                //  After Change the Document, we need to update the layer
                SwitchLayer(_SelectedLayerIndex); // 3 is the "normal" layer
            }

            return validDoc;
        }

        /// <summary>
        /// Returns the index of the currently selected layer
        /// </summary>
        public int GetSelectedLayerIndex()
        {
            return _SelectedLayerIndex;
        }

        /// <summary>
        /// Returns the name of the currently selected layer
        /// </summary>
        public string GetSelectedLayerName()
        {
            return LayerIndexToName(_SelectedLayerIndex);
        }

        /// <summary>
        /// Returns the number of master layers in the system
        /// </summary>
        public int GetMasterLayerCount()
        {
            return 5;
        }

        /// <summary>
        /// Change the internal selected layer of the binding object
        /// All future calls to layer related APIs will be use this new layer
        /// </summary>
        public bool SwitchLayer(int index)
        {
            bool hasLayer = false;
            string layerName = LayerIndexToName(index);

            if (layerName != "")
            {
                _SelectedLayer = Late.Invoke(_Document, "LayerByName", layerName);
                _SelectedLayerIndex = index;
                hasLayer = true;
            }

            return hasLayer;
        }

        /// <summary>
        /// Change the internal selected layer of the binding object
        /// All future calls to layer related APIs will be use this new layer
        /// </summary>
        public bool SwitchLayer(string name)
        {
            bool validName = false;
            int index = LayerNameToIndex(name.ToLower());

            if (index != -1)
            {
                SwitchLayer(index);
                validName = true;
            }

            return validName;
        }

        /// <summary>
        /// Return whether or not the current document is broadcasting a stream
        /// </summary>
        public bool IsBroadcasting()
        {
            int isBroadcasting = (int)Late.Invoke(_Document, "IsBroadcasting");
            return isBroadcasting == 1;
        }

        /// <summary>
        /// Toggles the Broadcast state
        /// If it is not broadcasting, start a broadcast
        /// If it is, stop it
        /// </summary>
        public void ToggleBroadcast()
        {
            if (IsBroadcasting())
            {
                StopBroadcast();
            }
            else
            {
                StartBroadcast();
            }
        }
        public void StartBroadcast()
        {
            Late.Invoke(_Document, "Broadcast", "start");
        }
        public void StopBroadcast()
        {
            Late.Invoke(_Document, "Broadcast", "stop");
        }

        /// <summary>
        /// Returns whether or not the current document is recording to disk
        /// </summary>
        public bool IsRecording()
        {
            int isRecording = (int)Late.Invoke(_Document, "IsArchivingToDisk");
            return isRecording == 1;
        }

        /// <summary>
        /// Toggles the Record to Disk state
        /// If it is not recording, start the recording
        /// If it is, stop it
        /// </summary>
        public void ToggleRecord()
        {
            if (IsRecording())
            {
                StopRecording();
            }
            else
            {
                StartRecording();
            }
        }
        public void StartRecording()
        {
            Late.Invoke(_Document, "ArchiveToDisk", "start");
        }
        public void StopRecording()
        {
            Late.Invoke(_Document, "ArchiveToDisk", "stop");
        }

        /// <summary>
        /// Returns the total umber of shots in the selected layer
        /// To change which layer to count the shots in, use switch layers
        /// To get the total number of shots in the document, use GetTotalShotCount
        /// </summary>
        public int GetShotCount()
        {
            int shotCount = (int)Late.Invoke(_SelectedLayer, "ShotCount");
            return shotCount;
        }

        /// <summary>
        /// Returns the total number of shots in the document
        /// </summary>
        public int GetTotalShotCount()
        {
            int totalLayers = GetMasterLayerCount();
            int totalShotCount = 0;
            for (int i = 0; i < totalLayers; i++)
            {
                totalShotCount += (int)Late.Invoke(i, "ShotCount");
            }

            return totalShotCount;
        }

        /// <summary>
        /// Returns the ShotID for associated with the Shot name
        /// The name is the text under the shot in the Shot bin
        /// </summary>
        public int GetShotIDByName(string name)
        {
            int shotID = (int)Late.Invoke(_SelectedLayer, "ShotIDByName", name, 2);
            return shotID;
        }

        /// <summary>
        /// Returns the ShotID for the index within the selected layer
        /// The index is Zero based and from left to right in the layer
        /// </summary>
        public int GetShotIDByIndex(int index)
        {
            int shotID = (int)Late.Invoke(_SelectedLayer, "ShotIDByIndex", index);
            return shotID;
        }

        /// <summary>
        /// Returns the Shot COM object for the index
        /// Usually would not need to call this method directly
        /// 
        /// Returns null if there are no shots assoicated witht he index
        /// </summary>
        public object GetShotWithIndex(int index)
        {
            object shot = null;
            int shotID = GetShotIDByIndex(index);

            if (shotID != 0)
            {
                shot = GetShotWithID(shotID);
            }

            return shot;
        }

        /// <summary>
        /// Returns the Shot COM object for the shotID
        /// Usually would not need to call this method directly
        /// 
        /// Returns null if there are no shots assoicated witht he shotID
        /// </summary>
        public object GetShotWithID(int shotID)
        {
            return Late.Invoke(_Document, "ShotByShotID", shotID);
        }

        /// <summary>
        /// Returns the Shot COM object for the name
        /// Usually would not need to call this method directly
        /// 
        /// Returns null if there are no shots assoicated witht he name
        /// </summary>
        public object GetShotWithName(string name)
        {
            int shotID = GetShotIDByName(name);
            return GetShotWithID(shotID);
        }

        /// <summary>
        /// Returns the name of the shot associated with the shotID
        /// If there are no shots associated with that ID, returns an empty string
        /// </summary>
        public string GetShotNameWithID(int shotID)
        {
            object shot = GetShotWithID(shotID);
            if (shot != null)
            {
                return (string)Late.Get(shot, "Name");
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// Returns the name of the shot at "index"
        /// If there are no shots associated with the index, returns an empty string
        /// </summary>
        public string GetShotNameWithIndex(int index)
        {
            int shotID = GetShotIDByIndex(index);
            return GetShotNameWithID(shotID);
        }

        /// <summary>
        /// Sets the name of the shot associated with the shot ID
        /// Returns true if the name was set successfully
        /// Returns false if there is no shot associated with the shot ID
        /// </summary>
        public bool SetShotNameWithID(int shotID, string newName)
        {
            bool isShotValid = false;
            object shot = GetShotWithID(shotID);

            if (shot != null)
            {
                Late.Set(shot, "Name", newName);
                isShotValid = true;
            }

            return isShotValid;
        }

        /// <summary>
        /// Returns true if the name was set successfully
        /// Returns false if there is no shot called "oldName"
        /// </summary>
        public bool SetShotNameWithName(string oldName, string newName)
        {
            bool isShotValid = false;
            object shot = GetShotWithName(oldName);
            if (shot != null)
            {
                Late.Set(shot, "Name", newName);
                isShotValid = true;
            }
            return isShotValid;
        }

        /// <summary>
        /// Returns true if the speed input string is a valid transition speed name
        /// Returns false if it is not
        /// </summary>
        private bool isSpeedValid(string speed)
        {
            bool isValid = (speed == "slowest" || speed == "slow" || speed == "normal" || speed == "faster" || speed == "fastest");
            return isValid;
        }

        /// <summary>
        /// Returns a string array of all the valid Transition Speeds
        /// </summary>
        public string[] GetValidTransitionSpeeds()
        {
            return new string[] { "slowest", "slow", "normal", "faster", "fastest" };
        }

        /// <summary>
        /// Returns the currently select Transition speed name
        /// 
        /// For a list of valid transition speeds, see GetValidTransitionSpeeds()
        /// </summary>
        public string GetTransitionSpeed()
        {
            string transSpeed = (string)Late.Get(_Document, "TransitionSpeed");
            return transSpeed;
        }

        /// <summary>
        /// Sets the Transition speed of the durrent document
        /// Returns true if the speed is valie
        /// Returns false if the speed is invalid
        /// 
        /// For a list of valid transition speeds, see GetValidTransitionSpeeds()
        /// </summary>
        public bool SetTransitionSpeed(string speed)
        {
            string lowerSpeed = speed.ToLower();
            bool isValidSpeed = isSpeedValid(lowerSpeed);
            if (isValidSpeed)
            {
                Late.Set(_Document, "TransitionSpeed", lowerSpeed);
            }

            return isValidSpeed;
        }

        /// <summary>
        /// Returns the current sate of AutoLive in the selected document
        /// </summary>
        public bool IsAutoLiveOn()
        {
            int autoLiveOn = (int)Late.Get(_Document, "AutoLive");
            return autoLiveOn == 1;
        }

        /// <summary>
        /// Toggles the AutoLive state
        /// If it is off, turn it on
        /// If it is on, turn it off
        /// </summary>
        public void ToggleAutoLive()
        {
            SetAutoLive(!IsAutoLiveOn());
        }
        public void SetAutoLive(bool on)
        {
            Late.Set(_Document, "AutoLive", on);
        }

        /// <summary>
        /// Returns true if the index maps to a Transition popup in Wirecast, false otherwise
        /// </summary>
        public bool isTransitionIndexValid(int index)
        {
            return (index == 1 || index == 2);
        }

        /// <summary>
        /// Returns the index of the currently active Transition popup
        /// A value of 0 represents the left most popup in the Wirecast UI
        /// </summary>
        public int GetActiveTransitionIndex()
        {
            int activeTransIndex = (int)Late.Get(_Document, "ActiveTransitionIndex");
            return activeTransIndex;
        }

        /// <summary>
        /// Set the Active Transition popup in Wirecast
        /// A value of 0 represents the left most popup in the Wirecast UI
        /// </summary>
        public bool SetActiveTransitionIndex(int index)
        {
            bool isIndexValid = isTransitionIndexValid(index);
            if (isIndexValid)
            {
                Late.Set(_Document, "ActiveTransitionIndex", index);
            }
            return isIndexValid;
        }

        /// <summary>
        /// Returns true if the Audio is muted tot he speakers
        /// </summary>
        public bool IsAudioMutedToSpeakers()
        {
            int audioMuted = (int)Late.Get(_Document, "AudioMutedToSpeaker");
            return audioMuted == 1;
        }
        public void ToggleAudioMutedToSpeakers()
        {
            bool audioMuted = IsAudioMutedToSpeakers();
            SetAudioMutedToSpeakers(!audioMuted);
        }
        public void SetAudioMutedToSpeakers(bool muted)
        {
            Late.Set(_Document, "AudioMutedToSpeaker", muted);
        }

        /// <summary>
        /// Takes a snapshot still image of the current output and saves it as a JPEG to the given path
        /// </summary>
        public void SaveSnapshot(string path)
        {
            Late.Invoke(_Document, "SaveSnapshot", path);
        }

        /// <summary>
        /// Remove the media asset at the given path from Wirecast
        /// The path is not the shot name, but the actual media location on disk
        /// </summary>
        public void RemoveMedia(string path)
        {
            Late.Invoke(_Document, "RemoveMedia", path);
        }

        /// <summary>
        /// Creates a new shot with the asset located in the given path and adds it to the currently selected layer
        /// </summary>
        public int AddShotWithMedia(string path)
        {
            int shotID = (int)Late.Invoke(_SelectedLayer, "AddShotWithMedia", path);
            return shotID;
        }

        /// <summary>
        /// Removes the shot with the given ID from the currently selected layer
        /// Does nothing if the shot ID is invalid or not associated with any shots in the currently selected layer
        /// </summary>
        public void RemoveShotWithID(int shotID)
        {
            Late.Invoke(_SelectedLayer, "RemoveShotByID", shotID);
        }

        /// <summary>
        /// Removes the shot with the given name from the currently selected layer
        /// Does nothing if the name is not associated with any shots in the currently selected layer
        /// </summary>
        public void RemoveShotWithName(string name)
        {
            int shotID = GetShotIDByName(name);
            RemoveShotWithID(shotID);
        }

        /// <summary>
        /// Makes the active shot of the selected layer go live
        /// </summary>
        public void Go()
        {
            Late.Invoke(_SelectedLayer, "Go");
        }

        /// <summary>
        /// Returns true if the currently selected layer is visible, false otherwise
        /// </summary>
        public bool IsLayerVisible()
        {
            int visible = (int)Late.Get(_SelectedLayer, "Visible");
            return visible == 1;
        }

        /// <summary>
        /// Toggles the selected layer's visibility
        /// </summary>
        public void ToggleLayerVisibility()
        {
            bool visible = IsLayerVisible();
            Late.Set(_SelectedLayer, "Visible", !visible);
        }

        /// <summary>
        /// Returns true if the shot ID is the ID of current active shot of the currently selected layer
        /// Returns false if the shot ID is invalid or not in the currently selected layer
        /// </summary>
        public bool IsActiveShot(int shotID)
        {
            int activeShotID = GetActiveShotID();
            return activeShotID == shotID;
        }

        /// <summary>
        /// Returns true if the name is the name of the current active shot of the currently selected layer
        /// Returns false if the name is not
        /// </summary>
        public bool IsActiveShot(string name)
        {
            string activeShotName = GetActiveShotName();
            return activeShotName == name;
        }

        /// <summary>
        /// Returns the shot ID of the active shot, of the currently selected layer
        /// The Active shot is equivilent to the shot the user has clicked
        /// It doesn't mean the shot that is currently live or in preview, though it is possible the active shot is live or in preview
        /// </summary>
        public int GetActiveShotID()
        {
            int shotID = (int)Late.Get(_SelectedLayer, "ActiveShotID");
            return shotID;
        }

        /// <summary>
        /// Sets the active shot of the currently selected layer
        /// The Active shot is equivilent to the shot the user has clicked, so this method is the same as when a user clicked the shot
        /// It doesn't mean the shot that is currently live or in preview, though it is possible the active shot is live or in preview
        /// </summary>
        public bool SetActiveShot(int shotID)
        {
            bool isShotValid = false;
            object shot = GetShotWithID(shotID);
            if (shot != null)
            {
                Late.Set(_SelectedLayer, "ActiveShotID", shotID);
            }
            return isShotValid;
        }
        public bool SetActiveShot(string name)
        {
            bool isShotValid = false;
            object shot = GetShotWithName(name);
            if (shot != null)
            {
                int shotID = GetShotIDByName(name);
                Late.Set(_SelectedLayer, "ActiveShotID", shotID);
            }
            return isShotValid;
        }

        /// <summary>
        /// Returns the shot info of the shot currently in preview, of the currently selected layer
        /// The shot in preview is equivilent to the active shot
        /// </summary>
        public int GetPreviewShotID()
        {
            int shotID = (int)Late.Invoke(_SelectedLayer, "PreviewShotID");
            return shotID;
        }
        public string GetPreviewShotName()
        {
            int shotID = GetPreviewShotID();
            return GetShotNameWithID(shotID);
        }
        public object GetPreviewShot()
        {
            int shotID = GetPreviewShotID();
            return GetShotWithID(shotID);
        }

        /// <summary>
        /// Returns the shot info of the shot currently live, in the currently selected layer
        /// </summary>
        public int GetLiveShotID()
        {
            int shotID = (int)Late.Invoke(_SelectedLayer, "LiveShotID");
            return shotID;
        }
        public string GetLiveShotName()
        {
            int shotID = GetLiveShotID();
            return GetShotNameWithID(shotID);
        }
        public object GetLiveShot()
        {
            int shotID = GetLiveShotID();
            return GetShotWithID(shotID);
        }

        /// <summary>
        /// Returns the name of the active shot in the currently selected layer
        /// </summary>
        public string GetActiveShotName()
        {
            int shotID = GetActiveShotID();
            return GetShotNameWithID(shotID);
        }

        /// <summary>
        /// Returns true if the shot is currently in preview
        /// false otherwise
        /// </summary>
        public bool IsShotIDInPreview(int shotID)
        {
            object shot = GetShotWithID(shotID);
            int result = (int)Late.Invoke(shot, "Preview");

            return result == 1;
        }
        public bool IsShotNameInPreview(string name)
        {
            object shot = GetShotWithName(name);
            int result = (int)Late.Invoke(shot, "Preview");

            return result == 1;
        }

        /// <summary>
        /// Returns true if the shot is currently live
        /// false otherwise
        /// </summary>
        public bool IsShotIDLive(int shotID)
        {
            object shot = GetShotWithID(shotID);
            int result = (int)Late.Invoke(shot, "Live");

            return result == 1;
        }
        public bool IsShotNameLive(string name)
        {
            object shot = GetShotWithName(name);
            int result = (int)Late.Invoke(shot, "Live");

            return result == 1;
        }
    }
}
