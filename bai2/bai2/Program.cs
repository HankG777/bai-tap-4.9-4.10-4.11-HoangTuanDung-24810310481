using bai2;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace LabEquipmentCheckout
{
    public class Checkout
    {
        public string MaPhieu { get; set; } = "";
        public string MaSinhVien { get; set; } = "";
        public string HoTen { get; set; } = "";
        public string ThietBi { get; set; } = "";
        public int SoLuong { get; set; } = 1;
        public DateTime NgayMuon { get; set; } = DateTime.Today;
        public DateTime HanTra { get; set; } = DateTime.Today.AddDays(7);
        public DateTime? NgayTraThucTe { get; set; }
        public string TrangThai { get; set; } = "Đang mượn";
    }

    internal static class Program
    {
        [STAThread] static void Main() { ApplicationConfiguration.Initialize(); Application.Run(new MainForm()); }
    }

    public class MainForm : Form
    {
        BindingList<Checkout> data = new();
        BindingSource source = new();
        ErrorProvider errors = new();
        TextBox txtMa = new(), txtMSSV = new(), txtTen = new(), txtTim = new();
        ComboBox cbThietBi = new(), cbTrangThai = new(), cbLoc = new();
        NumericUpDown numSL = new() { Minimum = 1, Maximum = 1000, Value = 1 };
        DateTimePicker dtMuon = new(), dtHan = new(), dtTra = new();
        CheckBox chkDaTra = new();
        DataGridView grid = new();
        ToolStripStatusLabel status = new();

        public MainForm()
        {
            Text = "Lab Equipment Checkout"; StartPosition = FormStartPosition.CenterScreen; Size = new Size(1250, 760); MinimumSize = new Size(1000, 620);
            errors.ContainerControl = this; BuildUI(); RefreshGrid();
        }
        void BuildUI()
        {
            var menu = new MenuStrip(); var file = new ToolStripMenuItem("Tệp"); file.DropDownItems.Add("Xuất CSV...", null, (_, __) => ExportCsv()); file.DropDownItems.Add("Thoát", null, (_, __) => Close()); menu.Items.Add(file); MainMenuStrip = menu; Controls.Add(menu);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) }; root.RowStyles.Add(new RowStyle(SizeType.Absolute, 260)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); Controls.Add(root); root.BringToFront();
            root.Controls.Add(BuildFormPanel(), 0, 0); root.Controls.Add(BuildGridPanel(), 0, 1); var strip = new StatusStrip(); strip.Items.Add(status); root.Controls.Add(strip, 0, 2);
        }
        Control BuildFormPanel()
        {
            var outer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(8), BorderStyle = BorderStyle.FixedSingle }; outer.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); outer.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            var f = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 3 }; for (int i = 0; i < 4; i++) f.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25)); for (int i = 0; i < 3; i++) f.RowStyles.Add(new RowStyle(SizeType.Percent, 33.33f));
            cbThietBi.DropDownStyle = ComboBoxStyle.DropDownList; cbThietBi.Items.AddRange(new object[] { "Laptop", "Máy chiếu", "Bộ Arduino", "Đồng hồ vạn năng", "Bộ dụng cụ điện" }); cbThietBi.SelectedIndex = 0;
            cbTrangThai.DropDownStyle = ComboBoxStyle.DropDownList; cbTrangThai.Items.AddRange(new object[] { "Đang mượn", "Đã trả", "Quá hạn" }); cbTrangThai.SelectedIndex = 0;
            dtMuon.Format = dtHan.Format = dtTra.Format = DateTimePickerFormat.Short; dtMuon.Value = DateTime.Today; dtHan.Value = DateTime.Today.AddDays(7);
            chkDaTra.Text = "Đã trả thiết bị"; chkDaTra.AutoSize = true; chkDaTra.CheckedChanged += (_, __) => { dtTra.Enabled = chkDaTra.Checked; if (chkDaTra.Checked) { dtTra.Value = DateTime.Today; cbTrangThai.Text = "Đã trả"; } else cbTrangThai.Text = "Đang mượn"; };
            AddField(f, "Mã phiếu *", txtMa, 0, 0); AddField(f, "Mã sinh viên *", txtMSSV, 1, 0); AddField(f, "Họ tên sinh viên", txtTen, 2, 0); AddField(f, "Thiết bị *", cbThietBi, 3, 0);
            AddField(f, "Số lượng *", numSL, 0, 1); AddField(f, "Ngày mượn", dtMuon, 1, 1); AddField(f, "Hạn trả", dtHan, 2, 1); AddField(f, "Ngày trả thực tế", dtTra, 3, 1);
            AddField(f, "Tình trạng", cbTrangThai, 0, 2); AddField(f, "Đánh dấu trả", chkDaTra, 1, 2); dtTra.Enabled = false;
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill }; buttons.Controls.Add(Btn("Thêm", (_, __) => AddItem())); buttons.Controls.Add(Btn("Cập nhật", (_, __) => UpdateItem())); buttons.Controls.Add(Btn("Xóa", (_, __) => DeleteItem())); buttons.Controls.Add(Btn("Làm mới", (_, __) => ClearForm()));
            outer.Controls.Add(f, 0, 0); outer.Controls.Add(buttons, 0, 1); return outer;
        }
        void AddField(TableLayoutPanel p, string label, Control c, int col, int row, int span = 1) { var a = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(4) }; a.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); a.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); a.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0); c.Dock = DockStyle.Fill; a.Controls.Add(c, 0, 1); p.Controls.Add(a, col, row); if (span > 1) p.SetColumnSpan(a, span); }
        Control BuildGridPanel()
        {
            var p = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(0, 8, 0, 0) }; p.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); p.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var filter = new FlowLayoutPanel { Dock = DockStyle.Fill }; filter.Controls.Add(new Label { Text = "Tìm mã SV/tên thiết bị:", AutoSize = true, Margin = new Padding(3, 9, 3, 3) }); txtTim.Width = 220; txtTim.TextChanged += (_, __) => RefreshGrid(); filter.Controls.Add(txtTim); filter.Controls.Add(new Label { Text = "Lọc:", AutoSize = true, Margin = new Padding(3, 9, 3, 3) });
            cbLoc.DropDownStyle = ComboBoxStyle.DropDownList; cbLoc.Items.AddRange(new object[] { "Tất cả", "Đang mượn", "Quá hạn", "Đã trả" }); cbLoc.SelectedIndex = 0; cbLoc.SelectedIndexChanged += (_, __) => RefreshGrid(); filter.Controls.Add(cbLoc);
            grid.Dock = DockStyle.Fill; grid.ReadOnly = true; grid.AllowUserToAddRows = false; grid.AutoGenerateColumns = false; grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Col("Mã phiếu", "MaPhieu"); Col("Mã SV", "MaSinhVien"); Col("Họ tên", "HoTen"); Col("Thiết bị", "ThietBi"); Col("SL", "SoLuong"); Col("Ngày mượn", "NgayMuon", "dd/MM/yyyy"); Col("Hạn trả", "HanTra", "dd/MM/yyyy"); Col("Ngày trả", "NgayTraThucTe", "dd/MM/yyyy"); Col("Trạng thái", "TrangThai");
            grid.SelectionChanged += (_, __) => LoadSelected(); p.Controls.Add(filter, 0, 0); p.Controls.Add(grid, 0, 1); return p;
        }
        void Col(string title, string prop, string format = null) { var c = new DataGridViewTextBoxColumn { HeaderText = title, DataPropertyName = prop }; if (format != null) c.DefaultCellStyle.Format = format; grid.Columns.Add(c); }
        Button Btn(string t, EventHandler e) { var b = new Button { Text = t, Width = 105, Height = 30 }; b.Click += e; return b; }
        Checkout Selected() => grid.CurrentRow?.DataBoundItem as Checkout;
        string StatusFor(Checkout x) { if (x.NgayTraThucTe.HasValue) return "Đã trả"; return x.HanTra.Date < DateTime.Today ? "Quá hạn" : "Đang mượn"; }
        bool ValidateInput()
        {
            errors.Clear(); bool ok = true; void Err(Control c, string m) { errors.SetError(c, m); ok = false; }
            var old = Selected();
            if (string.IsNullOrWhiteSpace(txtMa.Text)) Err(txtMa, "Nhập mã phiếu."); else if (data.Any(x => x.MaPhieu.Equals(txtMa.Text.Trim(), StringComparison.OrdinalIgnoreCase) && !ReferenceEquals(x, old))) Err(txtMa, "Mã phiếu đã tồn tại.");
            if (string.IsNullOrWhiteSpace(txtMSSV.Text)) Err(txtMSSV, "Nhập mã sinh viên."); if (string.IsNullOrWhiteSpace(cbThietBi.Text)) Err(cbThietBi, "Chọn thiết bị.");
            if (dtHan.Value.Date < dtMuon.Value.Date) Err(dtHan, "Hạn trả không được trước ngày mượn.");
            if (chkDaTra.Checked && dtTra.Value.Date < dtMuon.Value.Date) Err(dtTra, "Ngày trả không được trước ngày mượn.");
            return ok;
        }
        Checkout ReadForm() => new Checkout { MaPhieu = txtMa.Text.Trim(), MaSinhVien = txtMSSV.Text.Trim(), HoTen = txtTen.Text.Trim(), ThietBi = cbThietBi.Text, SoLuong = (int)numSL.Value, NgayMuon = dtMuon.Value.Date, HanTra = dtHan.Value.Date, NgayTraThucTe = chkDaTra.Checked ? dtTra.Value.Date : null, TrangThai = chkDaTra.Checked ? "Đã trả" : (dtHan.Value.Date < DateTime.Today ? "Quá hạn" : "Đang mượn") };
        void AddItem() { if (!ValidateInput()) return; data.Add(ReadForm()); RefreshGrid(); ClearForm(); }
        void UpdateItem() { var x = Selected(); if (x == null) { MessageBox.Show("Chọn phiếu cần cập nhật."); return; } if (!ValidateInput()) return; var v = ReadForm(); x.MaPhieu = v.MaPhieu; x.MaSinhVien = v.MaSinhVien; x.HoTen = v.HoTen; x.ThietBi = v.ThietBi; x.SoLuong = v.SoLuong; x.NgayMuon = v.NgayMuon; x.HanTra = v.HanTra; x.NgayTraThucTe = v.NgayTraThucTe; x.TrangThai = v.TrangThai; source.ResetBindings(false); RefreshGrid(); }
        void DeleteItem() { var x = Selected(); if (x == null) return; if (MessageBox.Show($"Xóa phiếu {x.MaPhieu}?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) { data.Remove(x); RefreshGrid(); ClearForm(); } }
        void ClearForm() { errors.Clear(); txtMa.Clear(); txtMSSV.Clear(); txtTen.Clear(); numSL.Value = 1; cbThietBi.SelectedIndex = 0; dtMuon.Value = DateTime.Today; dtHan.Value = DateTime.Today.AddDays(7); chkDaTra.Checked = false; dtTra.Value = DateTime.Today; cbTrangThai.SelectedIndex = 0; grid.ClearSelection(); }
        void LoadSelected() { var x = Selected(); if (x == null) return; txtMa.Text = x.MaPhieu; txtMSSV.Text = x.MaSinhVien; txtTen.Text = x.HoTen; cbThietBi.Text = x.ThietBi; numSL.Value = Math.Min(numSL.Maximum, Math.Max(numSL.Minimum, x.SoLuong)); dtMuon.Value = x.NgayMuon; dtHan.Value = x.HanTra; chkDaTra.Checked = x.NgayTraThucTe.HasValue; if (x.NgayTraThucTe.HasValue) dtTra.Value = x.NgayTraThucTe.Value; cbTrangThai.Text = StatusFor(x); }
        void RefreshGrid()
        {
            foreach (var x in data) x.TrangThai = StatusFor(x);
            var q = data.AsEnumerable(); string s = txtTim.Text.Trim(); if (s.Length > 0) q = q.Where(x => x.MaSinhVien.Contains(s, StringComparison.OrdinalIgnoreCase) || x.ThietBi.Contains(s, StringComparison.OrdinalIgnoreCase));
            if (cbLoc.SelectedItem is string st && st != "Tất cả") q = q.Where(x => x.TrangThai == st);
            var list = q.ToList(); source.DataSource = new BindingList<Checkout>(list); grid.DataSource = source;
            status.Text = $"Tổng số phiếu: {data.Count} | Phiếu quá hạn: {data.Count(x => StatusFor(x) == "Quá hạn")} | Đang hiển thị: {list.Count}";
        }
        void ExportCsv()
        {
            using var d = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "PhieuMuon.csv" }; if (d.ShowDialog() != DialogResult.OK) return;
            var q = data.AsEnumerable(); string s = txtTim.Text.Trim(); if (s.Length > 0) q = q.Where(x => x.MaSinhVien.Contains(s, StringComparison.OrdinalIgnoreCase) || x.ThietBi.Contains(s, StringComparison.OrdinalIgnoreCase)); if (cbLoc.SelectedItem is string st && st != "Tất cả") q = q.Where(x => x.TrangThai == st);
            var sb = new StringBuilder("MaPhieu,MaSinhVien,HoTen,ThietBi,SoLuong,NgayMuon,HanTra,NgayTraThucTe,TrangThai\r\n");
            foreach (var x in q) sb.AppendLine(string.Join(",", new[] { x.MaPhieu, x.MaSinhVien, x.HoTen, x.ThietBi, x.SoLuong.ToString(), x.NgayMuon.ToString("dd/MM/yyyy"), x.HanTra.ToString("dd/MM/yyyy"), x.NgayTraThucTe?.ToString("dd/MM/yyyy") ?? "", StatusFor(x) }.Select(Csv)));
            File.WriteAllText(d.FileName, "\uFEFF" + sb, Encoding.UTF8); MessageBox.Show("Đã xuất CSV.");
        }
        string Csv(string s) => "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
    }
}
