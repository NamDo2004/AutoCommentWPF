using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WpfApp1.Models;
using System.Text.Json;
namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            ExcelPackage.License.SetNonCommercialOrganization("My book");
            InitializeComponent();
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            // Ds UserInfo rong
            List<HocSinh> hs = new List<HocSinh>();
            try
            {
                // Mo file excel
                OpenFileDialog openFileDiaLog = new OpenFileDialog();
                openFileDiaLog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";

                if(openFileDiaLog.ShowDialog() == true)
                {
                    var package = new ExcelPackage(new FileInfo(openFileDiaLog.FileName));

                    // Neu mo file thanh cong lay sheet dau tien
                    var worksheet = package.Workbook.Worksheets[2];

                    if(worksheet != null && worksheet.Dimension != null)
                    {
                        // Duyet tu dong thu hai den dong cuoi cung
                        // Trong file excel du lieu bat dau tu dong thu 6
                        for (int i = worksheet.Dimension.Start.Row + 5; i <= worksheet.Dimension.End.Row; i++)
                        {
                            try
                            {
                                // Bien duyet cot
                                // Cot 1 la STT, ten , diem bat dau tu cot 2
                                int j = 2;
                                // Lay ho ten tuong ung vi tri [6, 2]
                                // Tang j sau khi thuc hien lenh
                                string TenHS = worksheet.Cells[i, j++].Text;
                                // Lay diem noi ung vi tri [6, 3]
                                double diem_noi = Convert.ToDouble(worksheet.Cells[i, j++].Text);
                                // Lay diem nghe vi tri [6, 4]
                                double diem_nghe = Convert.ToDouble(worksheet.Cells[i, j++].Text);
                                // Lay diem doc/viet vi tri [6, 5]
                                double diem_doc_viet = Convert.ToDouble(worksheet.Cells[i, j++].Text);
                                // Lay tong diem vi tri [6, 6]
                                double diem_tong = worksheet.Cells[i, j++].GetValue<double>();
                                // Lay muc dat duoc vi tri [6, 7]
                                char muc_dat = Convert.ToChar(worksheet.Cells[i, j++].Text);
                                // Lay nhan xet vi tri [6, 8]
                                string nhan_xet = worksheet.Cells[i, j++].Text;
                                

                                // Tao hoc_sinh tu doi tuong lay duoc
                                HocSinh hoc_sinh = new HocSinh(TenHS, diem_noi, diem_nghe, diem_doc_viet, diem_tong, muc_dat, nhan_xet);

                                // add UserInfo vao list
                                hs.Add(hoc_sinh);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Loi dong: "+i+", chi tiet: "+ex.Message);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Khong tim thay du lieu phu hop trong file Excel");
                    }
                }
                else
                {
                    MessageBox.Show("Ban chua chon file Excel");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: "+ex.Message);
            }
            dtgExcel.ItemsSource = hs;
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            List<HocSinh> danhSachHocSinh = dtgExcel.ItemsSource.Cast<HocSinh>().ToList();
            AutoNX(danhSachHocSinh);
            string filePath = "";
            // Tao save file dialog de luu file excel
            SaveFileDialog dialog = new SaveFileDialog();

            // Loc cac file dinh dang excel
            dialog.Filter = "Excel | *.xlsx | Excel 2003 | *.xls";

            // Neu moi file va chon duoc noi luu thanh cong se luu duong dan
            if (dialog.ShowDialog() == true)
            {
                filePath = dialog.FileName;
            }

            // Neu duong dan null hoac rong thi bao khong hop le va return ham
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Duong dan bao cao khong hop le");
                return;
            }

            try
            {
                using (ExcelPackage p = new ExcelPackage())
                {
                    // Dat ten nguoi tao file
                    p.Workbook.Properties.Author = "Nam Do";
                    // Dat tieu de file
                    p.Workbook.Properties.Title = "Bao cao thong ke";
                    // Tao mot sheet de lam viec tren do
                    p.Workbook.Worksheets.Add("Ktra sheet");
                    // Lay sheet vua add ra de thao tac
                    ExcelWorksheet ws = p.Workbook.Worksheets[0];

                    // Dat ten cho sheet
                    ws.Name = "Ktra sheet";
                    // Fontsize mac dinh cho ca sheet
                    ws.Cells.Style.Font.Size = 14;
                    // Font family mac dinh cho ca sheet
                    ws.Cells.Style.Font.Name = "Times New Roman";
                    // Tao danh sach cac column header
                    string[] arrColumnHeader =
                    {
                        "Họ tên",
                        "Nói",
                        "Nghe",
                        "Đọc & Viết",
                        "Tổng điểm",
                        "Mức đạt được",
                        "Nhận xét"
                    };
                    //Lay so luong cot dua vao header
                    var countColHeader = arrColumnHeader.Count();
                    // merge cac column lai tu column 1 den so column header
                    // gan gia tri cho cell vua merge la thong ke thong in User Kteam
                    ws.Cells[1, 1].Value = "Kết quả kiểm tra định kì";
                    ws.Cells[1, 1, 1, countColHeader].Merge = true;
                    // in dam
                    ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                    // can giua
                    ws.Cells[1, 1, 1, countColHeader].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    int colIndex = 1;
                    int rowIndex = 2;
                    /*
                    // Tao cac header tu column header da tao tu ben tren
                    foreach (var item in arrColumnHeader)
                    {
                        var cell = ws.Cells[rowIndex, colIndex];

                        // set mau gray
                        var fill = cell.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                        // can chinh cac border
                        var border = cell.Style.Border;
                        border.Bottom.Style =
                            border.Top.Style =
                            border.Left.Style =
                            border.Right.Style = ExcelBorderStyle.Thin;
                        // gan gia tri
                        cell.Value = item;

                        colIndex++;
                    }
                    */
                    // Lay ra danh sach UserInfo tu ItemSource vua DataGrid
                    List<HocSinh> hoc_sinh = dtgExcel.ItemsSource.Cast<HocSinh>().ToList();

                    // voi moi item trong danh sach se ghi tren mot dong
                    foreach (var item in hoc_sinh)
                    {
                        // bat dau ghi tu cot 1. Excel bat dau tu 1 khong phai tu 0
                        colIndex = 1;

                        // rowIndex tuong ung tung dong du lieu
                        rowIndex++;

                        // gan gia tri cho tung cell
                        ws.Cells[rowIndex, colIndex++].Value = item.HoTen;
                        ws.Cells[rowIndex, colIndex++].Value = item.D_Noi;
                        ws.Cells[rowIndex, colIndex++].Value = item.D_Nghe;
                        ws.Cells[rowIndex, colIndex++].Value = item.D_Doc_Viet;
                        ws.Cells[rowIndex, colIndex++].Value = item.Tong_Diem;
                        ws.Cells[rowIndex, colIndex++].Value = item.mucDatDuoc;
                        ws.Cells[rowIndex, colIndex++].Value = item.nhanXet;
                    }
                    // Luu file
                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(filePath, bin);

                }
                MessageBox.Show("Xuat file thanh cong");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: "+ex.Message);
            }

        }

        private void btnNhanXet_Click(object sender, EventArgs e)
        {
            AutoNX(dtgExcel.ItemsSource.Cast<HocSinh>().ToList());
        }

        private void AutoNX(List<HocSinh> danhSach)
        {
            // Doc file nhan_xet.json
            string filepath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "nhan_xet.json");
            if (!File.Exists(filepath))
            {
                MessageBox.Show("Khong tim thay file nhan_xet.json");
                return;
            }
            string jsonText = File.ReadAllText(filepath);
            DanhSachNX dsNhanXet = JsonSerializer.Deserialize<DanhSachNX>(jsonText);

            // Tao danh sach tam de khong trung lap
            List<string> nhanXetTotTemp = new List<string>(dsNhanXet.tot);
            List<string> nhanXetTeTemp = new List<string>(dsNhanXet.te);
            Random rnd = new Random();

            // Gan nhan xet cho tung hoc sinh
            foreach(var hs in danhSach)
            {
                if(hs.mucDatDuoc == 'T')
                {
                    if(nhanXetTotTemp.Count == 0)
                    {
                        // reset de random tiep
                        nhanXetTotTemp = new List<string>(dsNhanXet.tot);
                    }
                    int index = rnd.Next(nhanXetTotTemp.Count);
                    hs.nhanXet = nhanXetTotTemp[index];
                    nhanXetTotTemp.RemoveAt(index);
                }
                else if(hs.mucDatDuoc == 'H' && nhanXetTeTemp.Count > 0)
                {
                    if(hs.mucDatDuoc == 'H')
                    {
                        // reset
                        nhanXetTeTemp = new List<string>(dsNhanXet.te);
                    }
                    int index = rnd.Next(nhanXetTeTemp.Count);
                    hs.nhanXet = nhanXetTeTemp[index];
                    nhanXetTeTemp.RemoveAt(index);
                }
            }
        }

        public class UserInfo
        {
            public string Name { get; set; }
            public DateTime Birthday { get; set; }

        }
    }
}
