using bai3;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RepairServiceTracker
{
    public class RepairOrder
    {
        public string MaDon { get; set; } = "";
        public string TenKhachHang { get; set; } = "";
        public string SoDienThoai { get; set; } = "";
        public string LoaiThietBi { get; set; } = "";
        public string MoTaLoi { get; set; } = "";
        public DateTime NgayTiepNhan { get; set; } = DateTime.Today;
        public decimal ChiPhiDuKien { get; set; }
        public string TrangThai { get; set; } = "Mới tiếp nhận";
        public string DuongDanAnh { get; set; } = "";
    }

    internal static class Program
    {
        [STAThread] static void Main() { ApplicationConfiguration.Initialize(); Application.Run(new MainForm()); }
    }

    public class MainForm : Form
    {
        BindingList<RepairOrder> data = new();
        BindingSource source = new();
        ErrorProvider errors = new();
        TextBox txtMa = new(), txtTen = new(), txtSDT = new(), txtMoTa = new(), txtTim = new(), txtAnh = new();
        ComboBox cbLoai = new(), cbTrangThai = new(), cbLoc = new();
        DateTimePicker dtNhan = new();
        NumericUpDown numChiPhi = new() { Minimum = 0, Maximum = 1000000000, DecimalPlaces = 0, ThousandsSeparator = true, Increment = 50000 };
        PictureBox pic = new();
        DataGridView grid = new();
        ToolStripStatusLabel status = new();
        string selectedImage = "";

        public MainForm() { Text = "Repair Service Tracker"; StartPosition = FormStartPosition.CenterScreen; Size = new Size(1300, 800); MinimumSize = new Size(1050, 650); errors.ContainerControl = this; BuildUI(); RefreshGrid(); }
        void BuildUI()
        {
            var menu = new MenuStrip(); var file = new ToolStripMenuItem("Tệp"); file.DropDownItems.Add("Xuất CSV...", null, (_, __) => ExportCsv()); file.DropDownItems.Add("Thoát", null, (_, __) => Close()); menu.Items.Add(file); MainMenuStrip = menu; Controls.Add(menu);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) }; root.RowStyles.Add(new RowStyle(SizeType.Absolute, 300)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); Controls.Add(root); root.BringToFront(); root.Controls.Add(BuildFormPanel(), 0, 0); root.Controls.Add(BuildGridPanel(), 0, 1); var strip = new StatusStrip(); strip.Items.Add(status); root.Controls.Add(strip, 0, 2);
        }
        Control BuildFormPanel()
        {
            var outer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(8), BorderStyle = BorderStyle.FixedSingle }; outer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 76)); outer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 24));
            var left = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 3 }; for (int i = 0; i < 4; i++) left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); for (int i = 0; i < 3; i++) left.RowStyles.Add(new RowStyle(SizeType.Percent, 33.33f));
            cbLoai.DropDownStyle = ComboBoxStyle.DropDownList; cbLoai.Items.AddRange(new object[] { "Điện thoại", "Laptop", "Máy tính để bàn", "Máy tính bảng", "Máy in", "Thiết bị khác" }); cbLoai.SelectedIndex = 0;
            cbTrangThai.DropDownStyle = ComboBoxStyle.DropDownList; cbTrangThai.Items.AddRange(new object[] { "Mới tiếp nhận", "Đang sửa", "Chờ linh kiện", "Hoàn tất", "Đã hủy" }); cbTrangThai.SelectedIndex = 0;
            dtNhan.Format = DateTimePickerFormat.Short; dtNhan.Value = DateTime.Today;
            AddField(left, "Mã đơn *", txtMa, 0, 0); AddField(left, "Tên khách hàng *", txtTen, 1, 0); AddField(left, "Số điện thoại *", txtSDT, 2, 0); AddField(left, "Loại thiết bị", cbLoai, 3, 0);
            AddField(left, "Ngày tiếp nhận", dtNhan, 0, 1); AddField(left, "Chi phí dự kiến", numChiPhi, 1, 1); AddField(left, "Trạng thái", cbTrangThai, 2, 1); AddField(left, "Ảnh thiết bị", BuildImagePicker(), 3, 1);
            AddField(left, "Mô tả lỗi *", txtMoTa, 0, 2, 4);
            var right = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 }; right.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); right.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            pic.Dock = DockStyle.Fill; pic.BorderStyle = BorderStyle.FixedSingle; pic.SizeMode = PictureBoxSizeMode.Zoom; right.Controls.Add(pic, 0, 0);
            var bts = new FlowLayoutPanel { Dock = DockStyle.Fill }; bts.Controls.Add(Btn("Thêm", (_, __) => AddItem())); bts.Controls.Add(Btn("Cập nhật", (_, __) => UpdateItem())); bts.Controls.Add(Btn("Xóa", (_, __) => DeleteItem())); bts.Controls.Add(Btn("Làm mới", (_, __) => ClearForm())); right.Controls.Add(bts, 0, 1);
            outer.Controls.Add(left, 0, 0); outer.Controls.Add(right, 1, 0); return outer;
        }
        Control BuildImagePicker() { var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 }; p.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); p.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 38)); txtAnh.ReadOnly = true; txtAnh.Dock = DockStyle.Fill; var b = new Button { Text = "...", Dock = DockStyle.Fill }; b.Click += (_, __) => ChooseImage(); p.Controls.Add(txtAnh, 0, 0); p.Controls.Add(b, 1, 0); return p; }
        void AddField(TableLayoutPanel p, string label, Control c, int col, int row, int span = 1) { var a = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(3) }; a.RowStyles.Add(new RowStyle(SizeType.Absolute, 21)); a.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); a.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0); c.Dock = DockStyle.Fill; a.Controls.Add(c, 0, 1); p.Controls.Add(a, col, row); if (span > 1) p.SetColumnSpan(a, span); }
        Control BuildGridPanel()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(0, 8, 0, 0) }; p.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); p.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var filter = new FlowLayoutPanel { Dock = DockStyle.Fill }; filter.Controls.Add(new Label { Text = "Tìm mã đơn/tên/SĐT:", AutoSize = true, Margin = new Padding(3, 9, 3, 3) }); txtTim.Width = 250; txtTim.TextChanged += (_, __) => RefreshGrid(); filter.Controls.Add(txtTim); filter.Controls.Add(new Label { Text = "Trạng thái:", AutoSize = true, Margin = new Padding(3, 9, 3, 3) }); cbLoc.DropDownStyle = ComboBoxStyle.DropDownList; cbLoc.Items.AddRange(new object[] { "Tất cả", "Mới tiếp nhận", "Đang sửa", "Chờ linh kiện", "Hoàn tất", "Đã hủy" }); cbLoc.SelectedIndex = 0; cbLoc.SelectedIndexChanged += (_, __) => RefreshGrid(); filter.Controls.Add(cbLoc);
            grid.Dock = DockStyle.Fill; grid.ReadOnly = true; grid.AllowUserToAddRows = false; grid.AutoGenerateColumns = false; grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Col("Mã đơn", "MaDon"); Col("Khách hàng", "TenKhachHang"); Col("Số điện thoại", "SoDienThoai"); Col("Thiết bị", "LoaiThietBi"); Col("Mô tả lỗi", "MoTaLoi"); Col("Ngày nhận", "NgayTiepNhan", "dd/MM/yyyy"); Col("Chi phí dự kiến", "ChiPhiDuKien", "N0"); Col("Trạng thái", "TrangThai");
            grid.SelectionChanged += (_, __) => LoadSelected(); p.Controls.Add(filter, 0, 0); p.Controls.Add(grid, 0, 1); return p;
        }
        void Col(string t, string prop, string fmt = null) { var c = new DataGridViewTextBoxColumn { HeaderText = t, DataPropertyName = prop }; if (fmt != null) c.DefaultCellStyle.Format = fmt; grid.Columns.Add(c); }
        Button Btn(string t, EventHandler e) { var b = new Button { Text = t, Width = 83, Height = 28 }; b.Click += e; return b; }
        RepairOrder Selected() => grid.CurrentRow?.DataBoundItem as RepairOrder;
        bool ValidateInput()
        {
            errors.Clear(); bool ok = true; void Err(Control c, string m) { errors.SetError(c, m); ok = false; }
            var old = Selected();
            if (string.IsNullOrWhiteSpace(txtMa.Text)) Err(txtMa, "Nhập mã đơn."); else if (data.Any(x => x.MaDon.Equals(txtMa.Text.Trim(), StringComparison.OrdinalIgnoreCase) && !ReferenceEquals(x, old))) Err(txtMa, "Mã đơn đã tồn tại.");
            if (string.IsNullOrWhiteSpace(txtTen.Text)) Err(txtTen, "Nhập tên khách hàng.");
            if (!System.Text.RegularExpressions.Regex.IsMatch(txtSDT.Text.Trim(), @"^\+?\d{9,11}$")) Err(txtSDT, "Số điện thoại phải có 9–11 chữ số.");
            if (string.IsNullOrWhiteSpace(txtMoTa.Text)) Err(txtMoTa, "Mô tả lỗi là bắt buộc.");
            if (numChiPhi.Value < 0) Err(numChiPhi, "Chi phí không được âm.");
            return ok;
        }
        RepairOrder ReadForm() => new RepairOrder { MaDon = txtMa.Text.Trim(), TenKhachHang = txtTen.Text.Trim(), SoDienThoai = txtSDT.Text.Trim(), LoaiThietBi = cbLoai.Text, MoTaLoi = txtMoTa.Text.Trim(), NgayTiepNhan = dtNhan.Value.Date, ChiPhiDuKien = numChiPhi.Value, TrangThai = cbTrangThai.Text, DuongDanAnh = selectedImage };
        void AddItem() { if (!ValidateInput()) return; data.Add(ReadForm()); RefreshGrid(); ClearForm(); }
        void UpdateItem() { var x = Selected(); if (x == null) { MessageBox.Show("Chọn đơn cần cập nhật."); return; } if (!ValidateInput()) return; var v = ReadForm(); x.MaDon = v.MaDon; x.TenKhachHang = v.TenKhachHang; x.SoDienThoai = v.SoDienThoai; x.LoaiThietBi = v.LoaiThietBi; x.MoTaLoi = v.MoTaLoi; x.NgayTiepNhan = v.NgayTiepNhan; x.ChiPhiDuKien = v.ChiPhiDuKien; x.TrangThai = v.TrangThai; x.DuongDanAnh = v.DuongDanAnh; source.ResetBindings(false); RefreshGrid(); }
        void DeleteItem() { var x = Selected(); if (x == null) return; if (MessageBox.Show($"Xóa đơn {x.MaDon}?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) { data.Remove(x); RefreshGrid(); ClearForm(); } }
        void ChooseImage() { using var d = new OpenFileDialog { Filter = "Ảnh (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|Tất cả tệp (*.*)|*.*" }; if (d.ShowDialog() != DialogResult.OK) return; selectedImage = d.FileName; txtAnh.Text = selectedImage; ShowImage(selectedImage); }
        void ShowImage(string path) { try { if (pic.Image != null) { var old = pic.Image; pic.Image = null; old.Dispose(); } if (!string.IsNullOrWhiteSpace(path) && File.Exists(path)) { using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite); using var tmp = Image.FromStream(fs); pic.Image = new Bitmap(tmp); } } catch { pic.Image = null; } }
        void ClearForm() { errors.Clear(); txtMa.Clear(); txtTen.Clear(); txtSDT.Clear(); txtMoTa.Clear(); txtAnh.Clear(); numChiPhi.Value = 0; cbLoai.SelectedIndex = 0; cbTrangThai.SelectedIndex = 0; dtNhan.Value = DateTime.Today; selectedImage = ""; ShowImage(""); grid.ClearSelection(); }
        void LoadSelected() { var x = Selected(); if (x == null) return; txtMa.Text = x.MaDon; txtTen.Text = x.TenKhachHang; txtSDT.Text = x.SoDienThoai; txtMoTa.Text = x.MoTaLoi; cbLoai.Text = x.LoaiThietBi; cbTrangThai.Text = x.TrangThai; dtNhan.Value = x.NgayTiepNhan; numChiPhi.Value = Math.Min(numChiPhi.Maximum, Math.Max(numChiPhi.Minimum, x.ChiPhiDuKien)); selectedImage = x.DuongDanAnh; txtAnh.Text = selectedImage; ShowImage(selectedImage); }
        void RefreshGrid()
        {
            var q = data.AsEnumerable(); string s = txtTim.Text.Trim(); if (s.Length > 0) q = q.Where(x => x.MaDon.Contains(s, StringComparison.OrdinalIgnoreCase) || x.TenKhachHang.Contains(s, StringComparison.OrdinalIgnoreCase) || x.SoDienThoai.Contains(s, StringComparison.OrdinalIgnoreCase));
            if (cbLoc.SelectedItem is string st && st != "Tất cả") q = q.Where(x => x.TrangThai == st);
            var list = q.ToList(); source.DataSource = new BindingList<RepairOrder>(list); grid.DataSource = source;
            decimal total = data.Where(x => x.TrangThai != "Hoàn tất" && x.TrangThai != "Đã hủy").Sum(x => x.ChiPhiDuKien);
            status.Text = $"Đơn sau lọc: {list.Count} / Tổng đơn: {data.Count} | Tổng chi phí dự kiến chưa hoàn tất: {total:N0} VNĐ";
        }
        void ExportCsv()
        {
            using var d = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "DonSuaChua.csv" }; if (d.ShowDialog() != DialogResult.OK) return;
            var q = data.AsEnumerable(); string s = txtTim.Text.Trim(); if (s.Length > 0) q = q.Where(x => x.MaDon.Contains(s, StringComparison.OrdinalIgnoreCase) || x.TenKhachHang.Contains(s, StringComparison.OrdinalIgnoreCase) || x.SoDienThoai.Contains(s, StringComparison.OrdinalIgnoreCase)); if (cbLoc.SelectedItem is string st && st != "Tất cả") q = q.Where(x => x.TrangThai == st);
            var sb = new StringBuilder("MaDon,TenKhachHang,SoDienThoai,LoaiThietBi,MoTaLoi,NgayTiepNhan,ChiPhiDuKien,TrangThai,DuongDanAnh\r\n");
            foreach (var x in q) sb.AppendLine(string.Join(",", new[] { x.MaDon, x.TenKhachHang, x.SoDienThoai, x.LoaiThietBi, x.MoTaLoi, x.NgayTiepNhan.ToString("dd/MM/yyyy"), x.ChiPhiDuKien.ToString("N0"), x.TrangThai, x.DuongDanAnh }.Select(Csv)));
            File.WriteAllText(d.FileName, "\uFEFF" + sb, Encoding.UTF8); MessageBox.Show("Đã xuất CSV.");
        }
        string Csv(string s) => "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
        protected override void OnFormClosed(FormClosedEventArgs e) { if (pic.Image != null) pic.Image.Dispose(); errors.Dispose(); base.OnFormClosed(e); }
    }
}
