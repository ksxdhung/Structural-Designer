namespace Structural_Designer
{
    partial class Calculate : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Calculate()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Calculate));
            this.rbStructural = this.Factory.CreateRibbonTab();
            this.grGio = this.Factory.CreateRibbonGroup();
            this.btnNewGio = this.Factory.CreateRibbonButton();
            this.btnGiotinh = this.Factory.CreateRibbonButton();
            this.btnGiodong = this.Factory.CreateRibbonButton();
            this.btnDongdat = this.Factory.CreateRibbonButton();
            this.btnThuyetminhGio = this.Factory.CreateRibbonButton();
            this.grDam = this.Factory.CreateRibbonGroup();
            this.btnNewDam = this.Factory.CreateRibbonButton();
            this.btnOpenDam = this.Factory.CreateRibbonButton();
            this.btnThongsodam = this.Factory.CreateRibbonButton();
            this.btnTinhtoandam = this.Factory.CreateRibbonButton();
            this.btnBotriThepdam = this.Factory.CreateRibbonButton();
            this.btnThuyetminhdam = this.Factory.CreateRibbonButton();
            this.btnVeDam = this.Factory.CreateRibbonButton();
            this.grCot = this.Factory.CreateRibbonGroup();
            this.grVach = this.Factory.CreateRibbonGroup();
            this.btnBeamData = this.Factory.CreateRibbonButton();
            this.rbStructural.SuspendLayout();
            this.grGio.SuspendLayout();
            this.grDam.SuspendLayout();
            this.SuspendLayout();
            // 
            // rbStructural
            // 
            this.rbStructural.Groups.Add(this.grGio);
            this.rbStructural.Groups.Add(this.grDam);
            this.rbStructural.Groups.Add(this.grCot);
            this.rbStructural.Groups.Add(this.grVach);
            this.rbStructural.Label = "Tính toán Kết cấu";
            this.rbStructural.Name = "rbStructural";
            // 
            // grGio
            // 
            this.grGio.Items.Add(this.btnNewGio);
            this.grGio.Items.Add(this.btnGiotinh);
            this.grGio.Items.Add(this.btnGiodong);
            this.grGio.Items.Add(this.btnDongdat);
            this.grGio.Items.Add(this.btnThuyetminhGio);
            this.grGio.Label = "Tải trọng";
            this.grGio.Name = "grGio";
            // 
            // btnNewGio
            // 
            this.btnNewGio.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewGio.Image = ((System.Drawing.Image)(resources.GetObject("btnNewGio.Image")));
            this.btnNewGio.Label = "Công trình mới";
            this.btnNewGio.Name = "btnNewGio";
            this.btnNewGio.ShowImage = true;
            // 
            // btnGiotinh
            // 
            this.btnGiotinh.Label = "Gió tĩnh";
            this.btnGiotinh.Name = "btnGiotinh";
            // 
            // btnGiodong
            // 
            this.btnGiodong.Label = "Gió động";
            this.btnGiodong.Name = "btnGiodong";
            // 
            // btnDongdat
            // 
            this.btnDongdat.Label = "Động đất";
            this.btnDongdat.Name = "btnDongdat";
            // 
            // btnThuyetminhGio
            // 
            this.btnThuyetminhGio.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnThuyetminhGio.Image = ((System.Drawing.Image)(resources.GetObject("btnThuyetminhGio.Image")));
            this.btnThuyetminhGio.Label = "Thuyết minh";
            this.btnThuyetminhGio.Name = "btnThuyetminhGio";
            this.btnThuyetminhGio.ShowImage = true;
            // 
            // grDam
            // 
            this.grDam.Items.Add(this.btnNewDam);
            this.grDam.Items.Add(this.btnOpenDam);
            this.grDam.Items.Add(this.btnThongsodam);
            this.grDam.Items.Add(this.btnTinhtoandam);
            this.grDam.Items.Add(this.btnBotriThepdam);
            this.grDam.Items.Add(this.btnThuyetminhdam);
            this.grDam.Items.Add(this.btnVeDam);
            this.grDam.Items.Add(this.btnBeamData);
            this.grDam.Label = "Tính toán dầm";
            this.grDam.Name = "grDam";
            // 
            // btnNewDam
            // 
            this.btnNewDam.Image = ((System.Drawing.Image)(resources.GetObject("btnNewDam.Image")));
            this.btnNewDam.Label = "Công trình mới";
            this.btnNewDam.Name = "btnNewDam";
            this.btnNewDam.ShowImage = true;
            this.btnNewDam.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewDam_Click);
            // 
            // btnOpenDam
            // 
            this.btnOpenDam.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenDam.Image")));
            this.btnOpenDam.Label = "Mở công trình";
            this.btnOpenDam.Name = "btnOpenDam";
            this.btnOpenDam.ShowImage = true;
            this.btnOpenDam.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenDam_Click);
            // 
            // btnThongsodam
            // 
            this.btnThongsodam.Image = ((System.Drawing.Image)(resources.GetObject("btnThongsodam.Image")));
            this.btnThongsodam.Label = "Thông số";
            this.btnThongsodam.Name = "btnThongsodam";
            this.btnThongsodam.ShowImage = true;
            // 
            // btnTinhtoandam
            // 
            this.btnTinhtoandam.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTinhtoandam.Image = ((System.Drawing.Image)(resources.GetObject("btnTinhtoandam.Image")));
            this.btnTinhtoandam.Label = "Tính toán";
            this.btnTinhtoandam.Name = "btnTinhtoandam";
            this.btnTinhtoandam.ShowImage = true;
            // 
            // btnBotriThepdam
            // 
            this.btnBotriThepdam.Image = ((System.Drawing.Image)(resources.GetObject("btnBotriThepdam.Image")));
            this.btnBotriThepdam.Label = "Bố trí thép";
            this.btnBotriThepdam.Name = "btnBotriThepdam";
            this.btnBotriThepdam.ShowImage = true;
            // 
            // btnThuyetminhdam
            // 
            this.btnThuyetminhdam.Image = ((System.Drawing.Image)(resources.GetObject("btnThuyetminhdam.Image")));
            this.btnThuyetminhdam.Label = "Xuất thuyết minh";
            this.btnThuyetminhdam.Name = "btnThuyetminhdam";
            this.btnThuyetminhdam.ShowImage = true;
            // 
            // btnVeDam
            // 
            this.btnVeDam.Image = ((System.Drawing.Image)(resources.GetObject("btnVeDam.Image")));
            this.btnVeDam.Label = "Vẽ ACAD";
            this.btnVeDam.Name = "btnVeDam";
            this.btnVeDam.ShowImage = true;
            // 
            // grCot
            // 
            this.grCot.Label = "Tính toán cột";
            this.grCot.Name = "grCot";
            // 
            // grVach
            // 
            this.grVach.Label = "Tính toán vách";
            this.grVach.Name = "grVach";
            // 
            // btnBeamData
            // 
            this.btnBeamData.Label = "Lấy dữ liệu";
            this.btnBeamData.Name = "btnBeamData";
            this.btnBeamData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBeamData_Click);
            // 
            // Calculate
            // 
            this.Name = "Calculate";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.rbStructural);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Calculate_Load);
            this.rbStructural.ResumeLayout(false);
            this.rbStructural.PerformLayout();
            this.grGio.ResumeLayout(false);
            this.grGio.PerformLayout();
            this.grDam.ResumeLayout(false);
            this.grDam.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab rbStructural;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grGio;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewGio;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGiotinh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGiodong;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThuyetminhGio;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grDam;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grCot;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grVach;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDongdat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewDam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenDam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThongsodam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTinhtoandam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBotriThepdam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnThuyetminhdam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVeDam;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBeamData;
    }

    partial class ThisRibbonCollection
    {
        internal Calculate Calculate
        {
            get { return this.GetRibbon<Calculate>(); }
        }
    }
}
