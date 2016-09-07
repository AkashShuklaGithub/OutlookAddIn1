using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace OutlookAddIn1
{

    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class FormRegion : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public FormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
                    //fill the data
                    this.myFillData();
                    //initialize the components
                    this.InitializeComponent();
                    //fill the comboboxes
                    this.myMethod();
        }

        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        public ComboBox siteComboBox;
        public ComboBox floorComboBox;
        public Label label1;
        public Label label2;
        public Panel panel1;
        public static System.Collections.Generic.List<OutlookAddIn1.Site> siteList;
        public static System.Collections.Generic.List<OutlookAddIn1.Floor> floorList;
        public static System.Collections.Generic.List<OutlookAddIn1.Room> roomList;
        public static List<Button> listButton;
        public static listModel test;

        private void myFillData()
        {
            //deserialize the text file to populate entries

            System.IO.StreamReader streamReader;
            //to do
            //change the hard coded path to a dynamic runtime path
            string filePath = @"C:\Users\akash2\documents\visual studio 2015\Projects\OutlookAddIn1\OutlookAddIn1\TextFile1.txt";

            using (System.IO.FileStream fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Open))
            {
                using (streamReader = new System.IO.StreamReader(fileStream))
                {
                    test = Newtonsoft.Json.JsonConvert.DeserializeObject<listModel>(streamReader.ReadToEnd());

                    siteList = new System.Collections.Generic.List<OutlookAddIn1.Site>();
                    for (int i = 0; i < test.site.Count; i++)
                    {
                        siteList.Add(test.site[i]);
                    }
                }
            }
            streamReader.Close();
        }

        private void myMethod()
        {

           //site name list
            //populate data
            List<string> siteNameList = new List<string>();
            for (int i = 0; i < siteList.Count; i++)
            {
                siteNameList.Add(siteList[i].siteName);
            }
            //set the datasource of siteCombobox as siteList
            this.siteComboBox.DataSource = siteNameList;
            //in case the selection of site changes
            this.siteComboBox.SelectedIndexChanged += SiteComboBox_SelectedIndexChanged;


           //floor name list
            //populate data
            floorList = new List<Floor>();
            for (int i = 0; i < test.site[0].floor.Count; i++)
            {
                floorList.Add(test.site[0].floor[i]);
            }
            //populate combobox
            List<string> floorNameList = new List<string>();
            for (int i = 0; i < floorList.Count; i++)
            {
                floorNameList.Add(floorList[i].floorName);
            }
            //set datasource of floor combobox as floor names list
            floorComboBox.DataSource = floorNameList;
            //in case the selection of floor changes
            this.floorComboBox.SelectedIndexChanged += FloorComboBox_SelectedIndexChanged;

            //room name list
            //populate data
            roomList = new List<Room>();
            int t = test.site[0].floor[0].room.Count;
            for (int i = 0; i < t; i++)
            {
                roomList.Add(test.site[0].floor[0].room[i]);
            }
            //clear the buttons
            panel1.Controls.Clear();

            //create button list
            listButton = new List<System.Windows.Forms.Button>();

            for (int i = 0; i < t; i++)
            {
                listButton.Add(new System.Windows.Forms.Button());
            }



            //foreach button
            for (int index = 0; index < t; index++)
            {
                panel1.Controls.Add(listButton[index]);
                //populate buttons
                int scale = 2;
                listButton[index].Location = new System.Drawing.Point(roomList[index].locationX * scale, roomList[index].locationY * scale);
                listButton[index].Name = "button" + index;
                listButton[index].Size = new System.Drawing.Size(roomList[index].sizeX * scale, roomList[index].sizeY * scale);
                listButton[index].Tag = roomList[index].tag;
                listButton[index].Click += new System.EventHandler(OutlookAddIn1.ThisAddIn.button1_Click);
                listButton[index].BackColor = System.Drawing.Color.LightYellow;
                listButton[index].TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                listButton[index].Text = roomList[index].roomName;
                listButton[index].TextChanged += new EventHandler(ThisAddIn.decideButtonColor);
            }

        }

        private void FloorComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            //room name list
            //populate data
            roomList = new List<Room>();
            int siteComboBoxSelectedIndex=siteComboBox.SelectedIndex;
            int floorComboxSelectedIndex = floorComboBox.SelectedIndex;
            int t = test.site[siteComboBoxSelectedIndex].floor[floorComboxSelectedIndex].room.Count;
            for (int i = 0; i < t; i++)
            {
                roomList.Add(test.site[siteComboBoxSelectedIndex].floor[floorComboxSelectedIndex].room[i]);
            }
            //clear the buttons
            panel1.Controls.Clear();

            //create button list
            listButton = new List<System.Windows.Forms.Button>();
            
            for (int i = 0; i < t; i++)
            {
                listButton.Add(new System.Windows.Forms.Button());
            }
            //foreach button
            for (int i = 0; i < t; i++)
            {
                panel1.Controls.Add(listButton[i]);
                //populate buttons
                int scale = 2;
                listButton[i].Location = new System.Drawing.Point(roomList[i].locationX * scale, roomList[i].locationY * scale);
                listButton[i].Name = "button" + i;
                listButton[i].Size = new System.Drawing.Size(roomList[i].sizeX * scale, roomList[i].sizeY * scale);
                listButton[i].Tag = roomList[i].tag;
                listButton[i].Click += new System.EventHandler(OutlookAddIn1.ThisAddIn.button1_Click);
                listButton[i].BackColor = System.Drawing.Color.LightYellow;
                listButton[i].TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                listButton[i].Text = roomList[i].roomName;
                listButton[i].TextChanged += new EventHandler(ThisAddIn.decideButtonColor);
            }

        }

        private void SiteComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            //repopulate the floor name list
            int newSelectedIndex = 0;
            newSelectedIndex = siteComboBox.SelectedIndex;

            //floor name list
            //populate data
            floorList.Clear();
            for (int i = 0; i < test.site[newSelectedIndex].floor.Count; i++)
            {
                floorList.Add(test.site[newSelectedIndex].floor[i]);
            }
            //populate combobox
            List<string> floorNameList = new List<string>();
            for (int i = 0; i < floorList.Count; i++)
            {
                floorNameList.Add(floorList[i].floorName);
            }
            //set datasource of floor combobox as floor names list
            this.floorComboBox.DataSource = floorNameList;
        }


        public static int indice = 0;

        private void loadFreeBusy(object sender, EventArgs e)
        {
            if (indice <= (listButton.Count - 1))
            {
                listButton[indice].Text = listButton[indice].Text + "'";
                indice++;
            }
        }


        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.siteComboBox = new System.Windows.Forms.ComboBox();
            this.floorComboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // siteComboBox
            // 
            this.siteComboBox.BackColor = System.Drawing.Color.GhostWhite;
            this.siteComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.siteComboBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.siteComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.siteComboBox.Location = new System.Drawing.Point(72, 195);
            this.siteComboBox.Name = "siteComboBox";
            this.siteComboBox.Size = new System.Drawing.Size(121, 24);
            this.siteComboBox.TabIndex = 0;
            // 
            // floorComboBox
            // 
            this.floorComboBox.BackColor = System.Drawing.Color.GhostWhite;
            this.floorComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.floorComboBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.floorComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.floorComboBox.Location = new System.Drawing.Point(72, 312);
            this.floorComboBox.Name = "floorComboBox";
            this.floorComboBox.Size = new System.Drawing.Size(121, 24);
            this.floorComboBox.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(30, 198);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(36, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = ":Site";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(22, 315);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(44, 15);
            this.label2.TabIndex = 3;
            this.label2.Text = ":Floor";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.panel1.Location = new System.Drawing.Point(199, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1316, 740);
            this.panel1.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(53, 410);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(94, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "Load Free/Busy";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new EventHandler(loadFreeBusy);
            // 
            // FormRegion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.floorComboBox);
            this.Controls.Add(this.siteComboBox);
            this.Name = "FormRegion";
            this.Size = new System.Drawing.Size(1405, 756);
            this.FormRegionShowing += new System.EventHandler(this.FormRegion_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.FormRegion_FormRegionClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

       
        #endregion

        #region Form Region Designer generated code

        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            manifest.FormRegionName = "BookRoom";
            manifest.ShowReadingPane = false;

        }

        #endregion

        private Button button1;

        public partial class FormRegionFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public FormRegionFactory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                FormRegion.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.FormRegionFactory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                FormRegion form = new FormRegion(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }


    partial class WindowFormRegionCollection
    {
        internal FormRegion FormRegion
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(FormRegion))
                        return (FormRegion)item;
                }
                return null;
            }
        }
    }
}
