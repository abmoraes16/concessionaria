using System;
using System.IO;
using NetOffice.ExcelApi;
using System.Text.RegularExpressions;

namespace concessionaria_classes
{
    public class Carro
    {
        public Validacao validacao = new Validacao();
        Application ex = new Application();
        public string placa;
        public void CadastrarCarros()
        {
            string op1;
            int duplicado;
            Regex rgx = new Regex(@"^\S{3}\d{4}$");
            do{
                Console.WriteLine("\nCADASTRO DE CARROS: \n");
                do{
                    do{
                        Console.Write("Placa: ");
                        placa = Console.ReadLine();
                    }while(!rgx.IsMatch(placa));
                duplicado = PesquisaDocumento(placa);
                }while(duplicado!=0);

                ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Carro.xls");
                int cont=1;
                do{
                    cont++;
                }while(ex.Cells[cont,1].Value!=null);
                
                ex.Cells[cont,1].Value = placa;
                Console.Write("Marca: ");
                ex.Cells[cont,2].Value = Console.ReadLine();
                Console.Write("Modelo: ");
                ex.Cells[cont,3].Value = Console.ReadLine();
                Console.Write("Ano Modelo: ");
                ex.Cells[cont,4].Value = Console.ReadLine();
                Console.Write("Ano Fabricação: ");
                ex.Cells[cont,5].Value = Console.ReadLine();
                Console.Write("Preço: ");
                ex.Cells[cont,6].Value = Console.ReadLine();
                ex.Cells[cont,7].Value = 0;

                ex.ActiveWorkbook.Save();
                ex.Quit();
                
                do
                {
                    Console.Write("\nDeseja realizar um novo cadastro? (S ou N)");
                    op1 = Console.ReadLine();
                } while (op1!="S" && op1!="N" && op1!="s" && op1!="n");
            } while(op1=="S" || op1=="s");
        }      

        public int PesquisaDocumento(string placa)
        {
            int cont=1;

            ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Carro.xls");

            do{
                    if(ex.Cells[cont,1].Value.ToString() == placa){
                        Console.WriteLine("Placa já cadastrada! Seu Idiota!");
                        ex.Quit();

                        return 1;
                    }
                cont++;
            }while(ex.Cells[cont,1].Value!=null);
            ex.Quit();
            
            return 0;
        }        
    }
}