using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace RM
{
    public partial class frmPesquisar : Form
    {
        Variavel va = new Variavel();
        Conecta co = new Conecta();
        int rm;

        public frmPesquisar()
        {
            InitializeComponent();
        }

        public void Selecionar()
        {
            if (checkBox1.Checked == false)
            {
                if (comboBox1.Text == "")
                {
                    string msg = "ESCOLHA UMA OPÇÃO PARA PESQUISA";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                    comboBox1.Focus();
                }
                else if (textBox1.Text == "")
                {
                    string msg = "INFORME UM PARÂMETRO PARA PESQUISA";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                    textBox1.Focus();
                }
                else if (comboBox1.Text == "RM")
                {
                    va.Opcao = "RM";
                    va.Matricula = Convert.ToInt32(textBox1.Text);
                }
                else if (comboBox1.Text == "Nome")
                {
                    va.Opcao = "Nome";
                    va.Nome = textBox1.Text;
                }
                else if (comboBox1.Text == "RG")
                {
                    va.Opcao = "RG";
                    va.Registro = textBox1.Text;
                }
            }
            else
            {
                va.Opcao = "Intervalo";
                va.Matricula = Convert.ToInt32(textBox2.Text);
                va.MatriculaFinal = Convert.ToInt32(textBox3.Text);
            }
            dataGridView1.Rows.Clear();
            foreach (DataRow item in co.Selecionar(va).Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = item["rm"].GetHashCode();
                dataGridView1.Rows[n].Cells[1].Value = item["nome"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = item["rg"].ToString();
                if (item["spdoc"].ToString() == "Sim")
                {
                    dataGridView1.Rows[n].Cells[3].Value = true;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Selecionar();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                Selecionar();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                comboBox1.Text = "";
                comboBox1.Enabled = false;
                textBox1.Clear();
                textBox1.Enabled = false;
                textBox2.Enabled = true;
                textBox3.Enabled = true;
            }
            else
            {
                comboBox1.Enabled = true;
                textBox1.Enabled = true;
                textBox2.Enabled = false;
                textBox3.Enabled = false;
            }
        }

        private void frmPesquisar_Load(object sender, EventArgs e)
        {
            textBox2.Enabled = false;
            textBox3.Enabled = false;
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox1.Focus();
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                Selecionar();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            rm = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value);
            if (e.ColumnIndex == dataGridView1.Columns[3].Index)
            {
                dataGridView1.EndEdit();  //Stop editing of cell.


                if ((bool)dataGridView1.Rows[e.RowIndex].Cells[3].Value)
                {
                    va.Matricula = rm;
                    va.SpDoc = "Sim";
                    co.spdoc(va);
                    string msg = "PRONTUARIO CADASTRADO NO SPDOC";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                }
                else
                {
                    va.Matricula = rm;
                    va.SpDoc = "Não";
                    co.spdoc(va);
                    string msg = "PRONTUARIO NÃO CADASTRADO NO SPDOC";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                }
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                if (textBox3.Text != "")
                {
                    if (Convert.ToInt32(textBox3.Text) < Convert.ToInt32(textBox2.Text))
                    {
                        string msg = "O PRIMEIRO RM NÃO PODE SER MAIOR QUE O ÚLTIMO";
                        frmMensagem mg = new frmMensagem(msg);
                        mg.ShowDialog();
                        textBox2.Clear();
                        textBox2.Focus();
                    }
                }
            }
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            if (textBox3.Text != "")
            {
                if (Convert.ToInt32(textBox3.Text) < Convert.ToInt32(textBox2.Text))
                {
                    string msg = "O ÚLTIMO RM NÃO PODE SER MENOR QUE O PRIMEIRO";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                    textBox3.Clear();
                    textBox3.Focus();
                }
            }
        }
    }
}
