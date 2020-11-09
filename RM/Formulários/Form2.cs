using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using MySql.Data.MySqlClient;

namespace RM
{
    public partial class Form2 : Form
    {
        int aux;
        string caminho;
        DataTable dt = new System.Data.DataTable();
        Conecta co = new Conecta();
        Variavel va = new Variavel();

        public Form2(int mat, string cam)
        {
            InitializeComponent();
            caminho = cam;
        }

        public void Chamar()
        {
            va.Caminho = caminho;
            int linha = 1;

            int cont = co.Excel(va).Rows.Count;


            while (linha < cont)
            {
                va.Opcao = co.Excel(va).Rows[linha][0].ToString();
                aux = co.CompararRM(va).Rows.Count;
                progressBar1.Maximum = cont - 1;
                if (aux != 1)
                {
                    va.Matricula = Convert.ToInt32(co.Excel(va).Rows[linha][0].ToString());
                    va.Nome = co.Excel(va).Rows[linha][1].ToString();
                    va.Registro = co.Excel(va).Rows[linha][2].ToString();
                    co.cadastro(va);
                }
                linha = linha + 1;
                progressBar1.Value++;
            }
            label1.Text = "FINALIZADO";
            string msg = "RMs CADASTRADOS";
            frmMensagem mg = new frmMensagem(msg);
            mg.ShowDialog();
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            System.Threading.Thread tFormAguarde = new System.Threading.Thread(new System.Threading.ThreadStart(Chamar));
            tFormAguarde.Start();
        }
    }

}
