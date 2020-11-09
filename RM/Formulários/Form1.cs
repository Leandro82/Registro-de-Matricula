using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Deployment.Application;

namespace RM
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;
        SaveFileDialog salvarArquivo = new SaveFileDialog();
        OpenFileDialog arquivoExcel = new OpenFileDialog();
        Variavel va = new Variavel();
        Conecta co = new Conecta();

        public Form1()
        {
            InitializeComponent();
        }

        public void Fechar()
        {
            timer1.Interval = 2000;
            timer1.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            salvarArquivo.FileName = "RM";
            salvarArquivo.DefaultExt = "*.xls";
            salvarArquivo.Filter = "Todos os Aquivos do Excel (*.xls)|*.xls| Todos os arquivos (*.*)|*.*";
            
            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Name = "Planilha1";
                xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 3]].Merge();
                xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Cells[1, 1] = "Lista de RMs";
                xlWorkSheet.Cells[1, 2].ColumnWidth = 40.71;
                xlWorkSheet.Cells[1, 3].ColumnWidth = 25;
                xlWorkSheet.Cells[1, 1].Font.Size = 16;
                xlWorkSheet.Cells[2, 1] = "RM";
                xlWorkSheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Cells[2, 2] = "Nome";
                xlWorkSheet.Cells[2, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Cells[2, 3] = "RG";
                xlWorkSheet.Cells[2, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 3]].Font.Size = 12;

                new System.Threading.Thread(delegate()
                {
                    Export();
                }).Start();    
            }
            catch (Exception ex)
            {
                string msg = "Erro : " + ex.Message;
                frmMensagem mg = new frmMensagem(msg);
                mg.ShowDialog();
            }
        }

        private void Export()
        {
            System.Threading.Thread arquivo = new System.Threading.Thread(new System.Threading.ThreadStart(() =>
             {
                 if (salvarArquivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && salvarArquivo.FileName.Length > 0)
                 {
                     xlWorkBook.SaveAs(salvarArquivo.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                     Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                     xlWorkBook.Close(true, misValue, misValue);
                     xlApp.Quit();

                     liberarObjetos(xlWorkSheet);
                     liberarObjetos(xlWorkBook);
                     liberarObjetos(xlApp);

                     string msg = "O arquivo Excel foi criado com sucesso. Você pode encontrá-lo em : " + salvarArquivo.FileName;
                     frmMensagem mg = new frmMensagem(msg);
                     mg.ShowDialog();
                 }
             }));
            arquivo.SetApartmentState(System.Threading.ApartmentState.STA);
            arquivo.IsBackground = false;
            arquivo.Start();
        }

        private void liberarObjetos(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                string msg = "Ocorreu um erro durante a liberação do objeto " + ex.ToString();
                frmMensagem mg = new frmMensagem(msg);
                mg.ShowDialog();
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            new System.Threading.Thread(delegate()
            {
                Export1();
            }).Start();             
        }

        private void Export1()
        {
            System.Threading.Thread arquivo = new System.Threading.Thread(new System.Threading.ThreadStart(() =>
             {
                 if (arquivoExcel.ShowDialog() == DialogResult.OK)
                 {
                     string caminho = arquivoExcel.FileName;
                     int cont = co.Matricula().Rows.Count;
                     OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + arquivoExcel.FileName + ";Extended Properties=Excel 12.0;");
                     con.Open();
                     OleDbDataAdapter query = new OleDbDataAdapter(" SELECT * FROM [Planilha1$]", con);
                     DataTable dataTable = new DataTable();
                     query.Fill(dataTable);
                     int mat = dataTable.Rows.Count;
                     if (mat == 1)
                     {
                         string msg = "NÃO EXISTE DADOS PARA CADASTRAR";
                         frmMensagem mg = new frmMensagem(msg);
                         mg.ShowDialog();
                     }
                     else
                     {
                         Form2 form = new Form2(mat - 1, caminho);
                         form.ShowDialog();
                     }
                 }
             }));
            arquivo.SetApartmentState(System.Threading.ApartmentState.STA);
            arquivo.IsBackground = false;
            arquivo.Start();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var peq = new Form3("","","","");
            if (Application.OpenForms.OfType<Form3>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var peq = new frmPesquisar();
            if (Application.OpenForms.OfType<frmPesquisar>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            SendKeys.Send("{ESC}");
            timer1.Stop();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int l = 3;
            salvarArquivo.FileName = "RMs";
            salvarArquivo.DefaultExt = "*.xls";
            salvarArquivo.Filter = "Todos os Aquivos do Excel (*.xls)|*.xls| Todos os arquivos (*.*)|*.*";

            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 4]].Merge();
                xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Cells[1, 1] = "Lista de RMs";
                xlWorkSheet.Cells[1, 1].ColumnWidth = 10;
                xlWorkSheet.Cells[1, 2].ColumnWidth = 35;
                xlWorkSheet.Cells[1, 3].ColumnWidth = 20;
                xlWorkSheet.Cells[1, 1].Font.Size = 16;
                xlWorkSheet.Cells[2, 1] = "RM";
                xlWorkSheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Cells[2, 2] = "Nome";
                xlWorkSheet.Cells[2, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Cells[2, 3] = "RG";
                xlWorkSheet.Cells[2, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 4]].Font.Size = 12;
                int quant = co.Exportar().Rows.Count;
                progressBar1.Visible = true;
                progressBar1.Maximum = quant;
                foreach(DataRow item in co.Exportar().Rows)
                {
                    xlWorkSheet.Cells[l, 1] = item["rm"].ToString();
                    xlWorkSheet.Cells[l, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    xlWorkSheet.Cells[l, 2] = item["nome"].ToString();
                    xlWorkSheet.Cells[l, 3] = item["rg"].ToString();
                    xlWorkSheet.Cells[l, 3].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    l = l + 1;
                    progressBar1.Value++;
                }
                progressBar1.Value = 0;
                progressBar1.Visible = false;

                new System.Threading.Thread(delegate()
                {
                    Export();
                }).Start(); 
            }
            catch (Exception ex)
            {
                string msg = "Erro : " + ex.Message;
                frmMensagem mg = new frmMensagem(msg);
                mg.ShowDialog();
            }
        }
        
        private void button6_Click(object sender, EventArgs e)
        {
            
            int l = 3;
            int cont = co.Matricula().Rows.Count;
            int pos = 0;
            int[] num;
            num = new int[100000];

            progressBar1.Visible = true;
            progressBar1.Maximum = cont;


            for (int i = 0; i < cont; i++)
            {
                va.Matricula = i;
                int aux = co.ExportarBrancos(va).Rows.Count;

                if (aux == 0 && i != 0)
                {
                    num[pos] = i;
                    pos = pos + 1;
                }
                progressBar1.Value++;
            }
            
            progressBar1.Value = 0;

            salvarArquivo.FileName = "RMs em Branco";
            salvarArquivo.DefaultExt = "*.xls";
            salvarArquivo.Filter = "Todos os Aquivos do Excel (*.xls)|*.xls| Todos os arquivos (*.*)|*.*";

            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 4]].Merge();
                xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                xlWorkSheet.Cells[1, 1] = "Lista de RMs em branco";
                xlWorkSheet.Cells[1, 1].ColumnWidth = 10;
                xlWorkSheet.Cells[1, 1].Font.Size = 16;
                xlWorkSheet.Cells[2, 1] = "RM";
                xlWorkSheet.Cells[2, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                //progressBar1.Visible = true;
                progressBar1.Maximum = pos;
                for (int i = 0; i < pos; i++)
                {
                    xlWorkSheet.Cells[l, 1] = num[i];
                    xlWorkSheet.Cells[l, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    l = l + 1;
                    progressBar1.Value++;
                }
                progressBar1.Value = 0;
                progressBar1.Visible = false;

                new System.Threading.Thread(delegate()
                {
                    Export();
                }).Start(); 
            }
            catch (Exception ex)
            {
                string msg = "Erro : " + ex.Message;
                frmMensagem mg = new frmMensagem(msg);
                mg.ShowDialog();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Version appVersion;

            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                appVersion = ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            else
            {
                appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            }

            toolStripLabel2.Text = Convert.ToString(appVersion);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var peq = new Form3("", "", "", "");
            if (Application.OpenForms.OfType<Form3>().Count() > 0)
            {
                Application.OpenForms[peq.Name].Focus();
            }
            else
            {
                peq.Show();
            }
        }
    }
}
