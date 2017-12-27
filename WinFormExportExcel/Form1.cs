using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

//本例 必须添加引用：Microsoft.Office.Interop.Excel.dll （本例已包含）
//本例只是抛砖引玉，能实现基本功能，其他的请自行研究完善吧^^

namespace WinFormExportExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Bind();//从数据库取出数据
            //如果界面上不需要dataGridView控件，可以将dataGridView1控件隐藏掉！
            //dataGridView1.Visible = false;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Bind();//从数据库取出数据
        }
        /// <summary>
        /// 这里应该是从数据库中取出来的数据(Select的过程略)
        /// 将获得的DataTable添加到dataGridView的数据源中
        /// </summary>
        private void Bind()
        {
            //模拟数据：
            DataTable dt1 = new DataTable();//假如是从数据库取出来的数据
            dt1.Columns.Add("编号", typeof(string));
            dt1.Columns.Add("姓名", typeof(string));
            dt1.Columns.Add("性别", typeof(string));
            dt1.Columns.Add("年龄", typeof(string));

            dt1.Rows.Add("1", "天一", "女", "21");
            dt1.Rows.Add("2", "牛二", "男", "22");
            dt1.Rows.Add("3", "张三", "男", "20");
            dt1.Rows.Add("4", "李四", "女", "19");
            dt1.Rows.Add("5", "王五", "男", "25");
            dt1.Rows.Add("6", "赵六", "男", "24");
            dt1.Rows.Add("7", "田七", "男", "22");
            dt1.Rows.Add("8", "王八", "男", "21");
            dt1.Rows.Add("9", "白九", "男", "20");
            dt1.Rows.Add("10", "老十", "男", "24");
            dt1.Rows.Add("11", "石依", "女", "22");

            dataGridView1.DataSource = dt1;//添加到dataGridView的数据源中       
        }
        //导出Excel
        private void button1_Click(object sender, EventArgs e)
        {
            ClsExcel cExcel = null;
            cExcel = new ClsExcel();

            string fileName = "人员列表_"+DateTime.Now.ToString("yyyyMMddHHmmss"); //Excel文件名
            string sheetName = "全体人员";//sheet页的名称

            cExcel.ExportToExcel(ref dataGridView1, fileName, sheetName);
            cExcel.Close(false);
        }
        //我的腾讯微博^^：
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("http://t.qq.com/djk8888");
        }           
    }
}
