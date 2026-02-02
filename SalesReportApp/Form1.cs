using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using Microsoft.Reporting.WinForms;
using DataAccessLayer.Models;

namespace SalesReportApp
{
    public partial class Form1 : Form
    {
        // 1️⃣ Read connection string from App.config
        string connStr = ConfigurationManager
                            .ConnectionStrings["SalesDB"]
                            .ConnectionString;

        public Form1()
        {
            InitializeComponent();
            this.Load += Form1_Load;

            // 4️⃣ Add buttons programmatically
            AddReportButtons();
        }

        // 2️⃣ Load report when form opens
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // Set local processing
                reportViewer1.ProcessingMode = ProcessingMode.Local;

                // Set report path
                string reportPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "Reports",
                    "ProductSalesReport.rdlc"
                );

                if (!File.Exists(reportPath))
                {
                    MessageBox.Show($"Report file not found:\n{reportPath}");
                    return;
                }

                reportViewer1.LocalReport.ReportPath = reportPath;

                // Set professional page settings
                reportViewer1.SetDisplayMode(DisplayMode.PrintLayout); // show like printed page
                reportViewer1.ZoomMode = ZoomMode.PageWidth;            // scale to page width
                reportViewer1.LocalReport.EnableExternalImages = true;  // if using logos

                // Bind data
                ReportDataSource rds =
                    new ReportDataSource("ProductSalesDataSet", GetSales());

                reportViewer1.LocalReport.DataSources.Clear();
                reportViewer1.LocalReport.DataSources.Add(rds);

                // Refresh report
                reportViewer1.RefreshReport();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading report:\n{ex.Message}");
            }
        }

        // 3️⃣ Get data from database
        private List<ProductSales> GetSales()
        {
            var list = new List<ProductSales>();

            try
            {
                using (SqlConnection con = new SqlConnection(connStr))
                {
                    string query = "SELECT ProdCat, SubCat, OrderYear, OrderQtr, Sales, Region, Manager FROM ProductSales";


                    using (SqlCommand cmd = new SqlCommand(query, con))
                    {
                        con.Open();
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            while (dr.Read())
                            {
                                list.Add(new ProductSales
                                {
                                    ProdCat = dr["ProdCat"].ToString(),
                                    SubCat = dr["SubCat"].ToString(),
                                    OrderYear = Convert.ToInt32(dr["OrderYear"]),
                                    OrderQtr = dr["OrderQtr"].ToString(),
                                    Sales = Convert.ToDecimal(dr["Sales"]),
                                    Region = dr["Region"] != DBNull.Value ? dr["Region"].ToString() : "",
                                    Manager = dr["Manager"] != DBNull.Value ? dr["Manager"].ToString() : ""

                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error fetching sales data:\n{ex.Message}");
            }

            return list;
        }

        // 4️⃣ Add print/export buttons
        private void AddReportButtons()
        {
            // Print button
            Button btnPrint = new Button
            {
                Text = "Print",
                Top = 10,
                Left = 10,
                Width = 80
            };
            btnPrint.Click += (s, e) => PrintReport();
            this.Controls.Add(btnPrint);

            // Export PDF button
            Button btnPdf = new Button
            {
                Text = "Export PDF",
                Top = 10,
                Left = 100,
                Width = 100
            };
            btnPdf.Click += (s, e) => ExportReportToPdf();
            this.Controls.Add(btnPdf);

            // Export HTML button
            Button btnHtml = new Button
            {
                Text = "Export HTML",
                Top = 10,
                Left = 210,
                Width = 100
            };
            btnHtml.Click += (s, e) => ExportReportToHtml();
            this.Controls.Add(btnHtml);
        }

        // 5️⃣ Print
        private void PrintReport()
        {
            try
            {
                reportViewer1.PrintDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error printing report:\n{ex.Message}");
            }
        }

        // 6️⃣ Export to PDF
        private void ExportReportToPdf()
        {
            try
            {
                Warning[] warnings;
                string[] streamIds;
                string mimeType = string.Empty;
                string encoding = string.Empty;
                string extension = string.Empty;

                byte[] bytes = reportViewer1.LocalReport.Render(
                    "PDF",   // PDF format
                    null,
                    out mimeType,
                    out encoding,
                    out extension,
                    out streamIds,
                    out warnings);

                string pdfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ProductSalesReport.pdf");
                File.WriteAllBytes(pdfPath, bytes);

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = pdfPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to PDF: {ex.Message}");
            }
        }

        // 7️⃣ Export to HTML
        private void ExportReportToHtml()
        {
            try
            {
                Warning[] warnings;
                string[] streamIds;
                string mimeType = string.Empty;
                string encoding = string.Empty;
                string extension = string.Empty;

                byte[] bytes = reportViewer1.LocalReport.Render(
                    "HTML4.0",
                    null,
                    out mimeType,
                    out encoding,
                    out extension,
                    out streamIds,
                    out warnings);

                string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ProductSalesReport.html");
                File.WriteAllBytes(htmlPath, bytes);

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = htmlPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to HTML: {ex.Message}");
            }
        }
    }
}
