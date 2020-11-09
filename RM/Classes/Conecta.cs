using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data;
using System.Data.OleDb;

namespace RM
{
    class Conecta
    {
        public MySqlConnection conexao;

        public string Endereco()
        {
            StringConexao str = new StringConexao();
            return str.Endereco();
        }

        public void cadastro(Variavel va)
        {
            try
            {
                conexao = new MySqlConnection(Endereco());
                conexao.Open();
                string inserir = "INSERT INTO matricula(rm, nome, rg)VALUES('" + va.Matricula + "','" + va.Nome + "','" + va.Registro + "')";              
                MySqlCommand comandos = new MySqlCommand(inserir, conexao);
                comandos.ExecuteNonQuery();              
                conexao.Close();
            }
            catch(Exception ex) 
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public void Editar(Variavel va)
        {
            conexao = new MySqlConnection(Endereco());
            conexao.Open();
            string alterar = "UPDATE matricula SET rm = '" + va.Matricula + "',nome = '" + va.Nome + "',rg = '" + va.Registro + "'WHERE rm = '" + va.Matricula + "'";
            MySqlCommand comando = new MySqlCommand(alterar, conexao);
            comando.ExecuteNonQuery();
            conexao.Close();
        }

        public void Excluir(Variavel va)
        {
            conexao = new MySqlConnection(Endereco());
            conexao.Open();
            string alterar = "DELETE FROM matricula WHERE rm = '" + va.Matricula + "'";
            MySqlCommand comando = new MySqlCommand(alterar, conexao);
            comando.ExecuteNonQuery();
            conexao.Close();
        }

        public void spdoc(Variavel va)
        {
            try
            {
                conexao = new MySqlConnection(Endereco());
                conexao.Open();
                string atualizar = "update matricula set spdoc='" + va.SpDoc + "' where rm='" + va.Matricula + "'";
                MySqlCommand comandos = new MySqlCommand(atualizar, conexao);
                comandos.ExecuteNonQuery();
                conexao.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable Selecionar(Variavel va)
        {
            DataTable dt = new System.Data.DataTable();
            try
            {
                if (va.Opcao == "Nome")
                {
                    conexao = new MySqlConnection(Endereco());
                    conexao.Open();
                    string selecionar = "SELECT rm, nome, rg, spdoc FROM matricula where nome like'%" + va.Nome + "%'";
                    MySqlDataAdapter comandos = new MySqlDataAdapter(selecionar, conexao);
                    comandos.Fill(dt);
                    conexao.Close();
                }
                else if(va.Opcao == "RG")
                {
                    conexao = new MySqlConnection(Endereco());
                    conexao.Open();
                    string selecionar = "SELECT rm, nome, rg, spdoc FROM matricula where rg='" + va.Registro + "'";
                    MySqlDataAdapter comandos = new MySqlDataAdapter(selecionar, conexao);
                    comandos.Fill(dt);
                    conexao.Close();
                }
                else if (va.Opcao == "RM")
                {
                    conexao = new MySqlConnection(Endereco());
                    conexao.Open();
                    string selecionar = "SELECT rm, nome, rg, spdoc FROM matricula where rm='" + va.Matricula + "'";
                    MySqlDataAdapter comandos = new MySqlDataAdapter(selecionar, conexao);
                    comandos.Fill(dt);
                    conexao.Close();
                }
                else if (va.Opcao == "Intervalo")
                {
                    conexao = new MySqlConnection(Endereco());
                    conexao.Open();
                    string selecionar = "SELECT rm, nome, rg, spdoc FROM matricula where rm between'" + va.Matricula + "' and '" + va.MatriculaFinal + "'";
                    MySqlDataAdapter comandos = new MySqlDataAdapter(selecionar, conexao);
                    comandos.Fill(dt);
                    conexao.Close();
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable Exportar()
        {
            try
            {
                conexao = new MySqlConnection(Endereco());
                conexao.Open();
                string selecionar = "SELECT * FROM matricula ORDER BY rm";
                MySqlDataAdapter comandos = new MySqlDataAdapter(selecionar, conexao);
                DataTable dt = new System.Data.DataTable();
                comandos.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable ExportarBrancos(Variavel va)
        {
            try
            {
                conexao = new MySqlConnection(Endereco());
                conexao.Open();
                string selecionar = "SELECT rm FROM matricula WHERE rm = '" + va.Matricula + "'";
                MySqlDataAdapter comandos = new MySqlDataAdapter(selecionar, conexao);
                DataTable dt = new System.Data.DataTable();
                comandos.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable Excel(Variavel va)
        {
            try
            {
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + va.Caminho + ";Extended Properties=Excel 12.0;");
                con.Open();
                OleDbDataAdapter query = new OleDbDataAdapter(" SELECT * FROM [Planilha1$]", con);
                DataTable dt = new DataTable();
                query.Fill(dt);
                con.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable CompararRM(Variavel va)
        {
            try
            {
                conexao = new MySqlConnection(Endereco());
                conexao.Open();
                string selecionar = "SELECT rm FROM matricula WHERE rm = '" + va.Opcao + "'";
                MySqlDataAdapter comandos = new MySqlDataAdapter(selecionar, conexao);
                DataTable dt = new System.Data.DataTable();
                comandos.Fill(dt);
                conexao.Close();
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro de comandos: " + ex.Message);
            }
        }

        public DataTable Matricula()
        {
            string vSQL = "Select * FROM matricula";
            MySqlDataAdapter vDataAdapter = new MySqlDataAdapter(vSQL, Endereco());
            DataTable vTable = new DataTable();
            vDataAdapter.Fill(vTable);
            return vTable;
        }
    }
}

