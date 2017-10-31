using System;
using System.IO;
using NetOffice.ExcelApi;

namespace concessionaria_classes
{
    class Program
    {
        static Cliente cliente = new Cliente();
        static Venda venda = new Venda();
        static Carro carro = new Carro();
        
        static void Main(string[] args)
        {
            Menu();
        }

        static void Menu(){
            string op2;
            ValidarArquivos();
            
            do
            {
                Console.WriteLine("\nEscolha uma das opções abaixo\n1 - Cadastrar Clientes\n2 - Cadastrar Carros\n3 - Realizar Vendas\n4 - Vendas Realizadas\n\n0 - Sair");
                do
                {
                    op2 = Console.ReadLine();
                } while (op2 != "1" && op2 != "2" && op2 != "3" && op2 != "4" && op2 != "0");

                switch (op2)
                {
                    case "0": Environment.Exit(0); break;
                    case "1": cliente.CadastrarClientes(); break;
                    case "2": carro.CadastrarCarros(); break;
                    case "3": venda.RealizarVendas(); break;
                    case "4": venda.VendasDia(); break;
                }
            } while (op2 != "0");
        }
        static void ValidarArquivos(){
            if(!File.Exists(@"C:\Concessionaria\Cadastro_Cliente.xls")){
                Application ex = new Application();
                ex.Workbooks.Add();
                ex.Cells[1,1].Value = "DOCUMENTO";
                ex.Cells[1,2].Value = "NOME";
                ex.Cells[1,3].Value = "ENDERECO";
                ex.Cells[1,4].Value = "CIDADE";
                ex.Cells[1,5].Value = "ESTADO";
                ex.Cells[1,6].Value = "CEP";
                ex.ActiveWorkbook.SaveAs(@"C:\Concessionaria\Cadastro_Cliente.xls");
                ex.Quit();
                ex.Dispose();
            }

            if(!File.Exists(@"C:\Concessionaria\Cadastro_Carro.xls")){
                Application ex = new Application();
                ex.Workbooks.Add();
                ex.Cells[1,1].Value = "PLACA";
                ex.Cells[1,2].Value = "MARCA";
                ex.Cells[1,3].Value = "MODELO";
                ex.Cells[1,4].Value = "ANO_MODELO";
                ex.Cells[1,5].Value = "ANO_FABRICACAO";
                ex.Cells[1,6].Value = "PREÇO";
                ex.Cells[1,7].Value = "STATUS";
                ex.ActiveWorkbook.SaveAs(@"C:\Concessionaria\Cadastro_Carro.xls");
                ex.Quit();
            }

            if(!File.Exists(@"C:\Concessionaria\Cadastro_Venda.xls")){
                Application ex = new Application();
                ex.Workbooks.Add();
                ex.Cells[1,1].Value = "CLIENTE";
                ex.Cells[1,2].Value = "CARRO";
                ex.Cells[1,3].Value = "DATA";
                ex.Cells[1,4].Value = "FORMA DE PAGAMENTO";
                ex.Cells[1,5].Value = "PARCELAS";
                ex.Cells[1,6].Value = "VALOR";
                ex.ActiveWorkbook.SaveAs(@"C:\Concessionaria\Cadastro_Venda.xls");
                ex.Quit();            
            }
        }
    }
}
