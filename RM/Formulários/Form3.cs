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
    public partial class Form3 : Form
    {
        Conecta co = new Conecta();
        Variavel va = new Variavel();
        int cont;
        string ind1;
        public Form3(string rm, string nome, string rg, string ind)
        {
            InitializeComponent();
            textBox1.Text = rm;
            textBox2.Text = nome;
            textBox3.Text = rg;
            ind1 = ind;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int aux = 0;
            cont = co.Matricula().Rows.Count;

            Conecta cl = new Conecta();
            Variavel va = new Variavel();

            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "")
            {
                string msg = "PREENCHER TODOS OS CAMPOS";
                frmMensagem mg = new frmMensagem(msg);
                mg.ShowDialog();
            }
            else
            {
                if (cont == 0)
                {
                    va.Matricula = Convert.ToInt32(textBox1.Text);
                    va.Nome = textBox2.Text;
                    va.Registro = textBox3.Text;
                    cl.cadastro(va);
                    string msg = "RM CADASTRADO";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                }

                while (aux < cont)
                {
                    if (co.Matricula().Rows[aux]["rm"].ToString() != textBox1.Text || co.Matricula().Rows[aux]["rg"].ToString() == textBox3.Text)
                    {
                        va.Matricula = Convert.ToInt32(textBox1.Text);
                        va.Nome = textBox2.Text;
                        va.Registro = textBox3.Text;
                        cl.cadastro(va);
                        string msg = "RM CADASTRADO";
                        frmMensagem mg = new frmMensagem(msg);
                        mg.ShowDialog();
                        aux = cont;
                    }
                    else
                    {
                        string msg = "RM JÁ CADASTRADO";
                        frmMensagem mg = new frmMensagem(msg);
                        mg.ShowDialog();
                    }
                    aux = aux + 1;
                }
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Text == "Salvar")
            {
                if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "")
                {
                    string msg = "NÃO EXISTE EVENTO PARA EDITAR!!";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                }
                else
                {
                    va.Matricula = Convert.ToInt32(textBox1.Text);
                    va.Nome = textBox2.Text;
                    va.Registro = textBox3.Text;
                    co.Editar(va);
                    label5.Visible = false;
                    textBox1.BackColor = Color.White;
                    button1.Enabled = true;
                    button2.Text = "Editar/Excluir";
                    button3.Visible = false;
                    string msg = "REGISTRO ALTERADO!!";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                }
            }
            else
            {
                button1.Enabled = false;
                button2.Text = "Salvar";
                button3.Visible = true;
                textBox1.BackColor = Color.Yellow;
                label5.Visible = true;
                label5.Text = "<= Digite o RM e dê enter";
                textBox1.Focus();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                string msg = "NÃO EXISTE EVENTO PARA EXCLUIR";
                frmMensagem mg = new frmMensagem(msg);
                mg.ShowDialog();
            }
            else
            {
                va.Matricula = Convert.ToInt32(textBox1.Text);
                string message = "Deseja realmente excluir este evento?";
                string caption = "Confirmar exclusão";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result;

                result = MessageBox.Show(message, caption, buttons);

                if (result == System.Windows.Forms.DialogResult.No)
                {
                    this.Close();
                }
                else
                {
                    co.Excluir(va);
                    label5.Visible = false;
                    textBox1.BackColor = Color.White;
                    button1.Enabled = true;
                    button2.Text = "Editar/Excluir";
                    button3.Visible = false;
                    string msg = "REGISTRO EXCLUÍDO!!";
                    frmMensagem mg = new frmMensagem(msg);
                    mg.ShowDialog();
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            SendKeys.Send("{ESC}");
            timer1.Stop();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            label5.Visible = false;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox1.BackColor == Color.Yellow)
                {
                    va.Matricula = Convert.ToInt32(textBox1.Text);
                    va.Opcao = "RM";
                    int cont = co.Selecionar(va).Rows.Count;

                    if (cont > 0)
                    {
                        textBox2.Text = co.Selecionar(va).Rows[0]["nome"].ToString();
                        textBox3.Text = co.Selecionar(va).Rows[0]["rg"].ToString();
                    }
                    else
                    {
                        string msg = "NÃO EXISTE ALUNO CADASTRADO COM ESSE RM!!";
                        frmMensagem mg = new frmMensagem(msg);
                        mg.ShowDialog();
                    }
                }
            }
        }
    }
}
