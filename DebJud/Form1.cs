using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using System.Data.OleDb;
//using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using System.Globalization;
using Spire.Xls;

namespace DebJud
{
    public partial class Form1 : Form
    {
        private string caminhoPlanilhaSaldo;
        private string caminhoPlanilhaTarifa;
        private DataSet mDataSet;
        private MySqlConnection mConn;
        bool existeDiretorio;
        int inicioPlanilhaSaldo;
        int inicioPlanilhaTarifa;

        public Form1()
        {
            InitializeComponent();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }



        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                caminhoPlanilhaSaldo = textBox1.Text;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog2.FileName;
                caminhoPlanilhaTarifa = textBox2.Text;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //    try
            //    {
            //Conexão com banco
            mDataSet = new DataSet();
            mConn = new MySqlConnection("Server=10.11.17.30;Database=4030;Uid=root;Pwd=chinchila@acida12244819;");
            //mConn = new MySqlConnection("Server=192.168.0.107;Database=tarifa;Uid=Denis;Pwd=Dtf@4030;");
            mConn.Open();

            if (caminhoPlanilhaSaldo == null || caminhoPlanilhaTarifa == null)
                MessageBox.Show("Selecione os caminhos dos arquivos");

            if (existeDiretorio == false)
                System.IO.Directory.CreateDirectory("c:/RCArquivos");

            if (existeDiretorio == false)
                System.IO.Directory.CreateDirectory("c:/RCArquivos/auxiliar");


            string descricao = "";
            string codigoMov = "";

            if (radioButton1.Checked)
            {
                descricao = "REF. AO PGTO DE IPVA, PROTOCOLO";
                codigoMov = "61";
            }
            else
            {
                descricao = "REF. AO PGTO DE PROTESTO, PROTOCOLO";
                codigoMov = "60";
            }

            StreamWriter writePrincipal = new StreamWriter(@"C:\RCArquivos\Tarifas-" + DateTime.Now.Date.ToString("yyyyMMdd") + ".txt");
            writePrincipal.WriteLine("0175640300001562SICOOBDIVI" + codigoMov + "03" + DateTime.Now.Date.ToString("yyyyMMdd") + "                                                                                                                                                                  ");


            Conexao conectaPlanilha = new Conexao();
            DataSet output = conectaPlanilha.importarExcel(caminhoPlanilhaSaldo, caminhoPlanilhaTarifa);

            var wb = new XLWorkbook();//Varíavel para a planilha(Pasta de trabalho)
            var ws = wb.Worksheets.Add("DADOS");//Variavel para a planilha dentro do workbook

            ws.Cell("A1").Value = "Conta";
            ws.Cell("B1").Value = "Valor";
            ws.Cell("C1").Value = "Protocolo";

            dataGridView1.DataSource = output.Tables[0];
            dataGridView2.DataSource = output.Tables[1];

            for (int j = 0; j <= output.Tables[0].Rows.Count; j++)
            {
                if (output.Tables[0].Rows[j]["F1"].ToString().Length > 0)
                {
                    if (!output.Tables[0].Rows[j]["F1"].ToString().Contains("Conta"))
                    {
                        inicioPlanilhaSaldo = j;
                        break;
                    }
                }
            }

            for (int i = 0; i <= output.Tables[1].Rows.Count; i++)
            {
                string nomeColuna1 = dataGridView2.Columns[0].HeaderText;

                if (output.Tables[1].Rows[i][nomeColuna1].ToString().Length > 0)
                {
                    if (!output.Tables[1].Rows[i][nomeColuna1].ToString().Contains("Conta"))
                    {
                        inicioPlanilhaTarifa = i;
                        break;
                    }
                }
            }


            string nomeColunaContaArquivo = dataGridView2.Columns[0].HeaderText;
            string nomeColunaValorArquivo = dataGridView2.Columns[1].HeaderText;
            string nomeColunaProtocoloArquivo = dataGridView2.Columns[2].HeaderText;

            string nomeColunaContaSaldo = dataGridView1.Columns[0].HeaderText;
            string nomeColunaValorSaldoDisponivel = dataGridView1.Columns[25].HeaderText;
            string nomeColunaValorLimiteDisponivel = dataGridView1.Columns[34].HeaderText;

            int linha = 2;

            for (int i = inicioPlanilhaTarifa; i < output.Tables[1].Rows.Count; i++)
            {
                string contaArquivo = output.Tables[1].Rows[i][nomeColunaContaArquivo].ToString();
                double valorArquivo = Convert.ToDouble(output.Tables[1].Rows[i][nomeColunaValorArquivo]);
                int ProtocoloArquivo = Convert.ToInt32(output.Tables[1].Rows[i][nomeColunaProtocoloArquivo]);

                for (int j = inicioPlanilhaSaldo; j < output.Tables[0].Rows.Count; j++)
                {
                    string contaSaldo = output.Tables[0].Rows[j][nomeColunaContaSaldo].ToString();



                    if (contaArquivo == contaSaldo)
                    {
                        
                        double valorSaldoDisponivel = Convert.ToDouble(output.Tables[0].Rows[j][nomeColunaValorSaldoDisponivel]);
                        double valorLimiteDisponivel = Convert.ToDouble(output.Tables[0].Rows[j][nomeColunaValorLimiteDisponivel]);
                        double saldo = Math.Round(valorLimiteDisponivel + valorSaldoDisponivel, 2);

                        if (saldo >= valorArquivo)
                        {
                            writePrincipal.WriteLine("1D" + contaSaldo.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + valorArquivo.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                    "          000N" + descricao.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E").Replace("Ô", "O") + "                  ");

                        }
                        else if (saldo > 0)
                        {
                            double novoValor = valorArquivo - saldo;
                            ws.Cell("A" + linha.ToString()).Value = novoValor;

                            writePrincipal.WriteLine("1D" + contaSaldo.Replace("-", "").Replace(".", "").PadLeft(10, '0') + "                                                                  000                              " + saldo.ToString("N2").Replace(".", "").Replace(",", "").PadLeft(17, '0') +
                                                                   "          000N" + descricao.PadLeft(40, ' ').Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Ã", "A").Replace("Õ", "O").Replace("Í", "I").Replace("Ó", "O").Replace("Ê", "E").Replace("Ô", "O") + "                  ");

                            ws.Cell("A" + linha.ToString()).Value = contaSaldo;
                            ws.Cell("B" + linha.ToString()).Value = novoValor ;
                            ws.Cell("C" + linha.ToString()).Value = ProtocoloArquivo;
                            linha++;
                        }
                        else if (saldo <= 0)
                        {
                            ws.Cell("A" + linha.ToString()).Value = contaSaldo;
                            ws.Cell("B" + linha.ToString()).Value = valorArquivo;
                            ws.Cell("C" + linha.ToString()).Value = ProtocoloArquivo;
                            linha++;
                        }

                    }
                }



                //Liberar objetos da memoria
                ws.Dispose();
                wb.Dispose();

                //Salvar o arquivo no disco
                wb.SaveAs(@"c:\RCArquivos\auxiliar\" + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + ".xlsx");

                Workbook teste = new Workbook();
                teste.LoadFromFile(@"c:\RCArquivos\auxiliar\" + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + ".xlsx");
                teste.SaveToFile(@"C:\RCArquivos\Convertido"+ DateTime.Now.Minute.ToString() + ".xls", ExcelVersion.Version97to2003);


            }


            writePrincipal.Dispose();
            writePrincipal.Close();
            //}
            // catch (Exception exception)
            // {
            //    Console.WriteLine(exception);
            //     throw;
            //   }
            MessageBox.Show("Arquivo Gerado com sucesso!!");
        }



    }
}
