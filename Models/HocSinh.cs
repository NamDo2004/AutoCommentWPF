using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Models
{
    internal class HocSinh
    {
        private string hoTen;
        private double d_noi;
        private double d_nghe;
        private double d_doc_viet;
        private double tong_diem;
        private char MucDatDuoc;
        private string NhanXet;


        public HocSinh() { }
        public HocSinh(String hoTen, double d_noi, double d_nghe, double d_doc_viet, double tong_diem, char MucDatDuoc, string NhanXet)
        {
            this.hoTen = hoTen;
            this.d_noi = d_noi;
            this.d_nghe = d_nghe;
            this.d_doc_viet = d_doc_viet;
            this.tong_diem = tong_diem;
            this.MucDatDuoc = MucDatDuoc;
            this.NhanXet = NhanXet;
        }

        public String HoTen
        {
            get { return hoTen; }
            set { hoTen = value; }
        }

        public double D_Noi
        {
            get { return d_noi; }
            set { d_noi = value; }
        }

        public double D_Nghe
        {
            get { return d_nghe; }
            set { d_nghe = value; }
        }

        public double D_Doc_Viet
        {
            get { return d_doc_viet; }
            set { d_doc_viet = value; }
        }

        public double Tong_Diem
        {
            get { return tong_diem; }
            set { tong_diem = value; }
        }

        public char mucDatDuoc
        {
            get { return MucDatDuoc; }
            set { MucDatDuoc = value; }
        }

        public string nhanXet
        {
            get { return NhanXet; }
            set { NhanXet = value; }
        }
    }
}
