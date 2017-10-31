using System;
using System.IO;
using NetOffice.ExcelApi;
using System.Text.RegularExpressions;

namespace concessionaria_classes
{
    
    public class Venda
    {
        Application ex = new Application();
        Regex rgx = new Regex(@"^\S{3}\d{4}$");
        Cliente cliente = new Cliente();
        string formaPagamento;

            public void VendasDia(){

            string op1,data;



            do{

                Console.Write("Digite o dia para pesquisa (DD/MM/AAAA): ");

                data = Console.ReadLine();

                ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Cliente.xls");

                

                int cont=1;

                do{

                    if(ex.Cells[cont,3].Value.ToString() == data){

                        Console.WriteLine(ex.Cells[cont,1].Value.ToString()+"\t");

                        Console.WriteLine(ex.Cells[cont,2].Value.ToString()+"\t");

                        Console.WriteLine(ex.Cells[cont,3].Value.ToString()+"\t");

                        Console.WriteLine(ex.Cells[cont,4].Value.ToString()+"\t");

                        if(ex.Cells[cont,4].Value.ToString() == "VISTA"){

                            Console.WriteLine(ex.Cells[cont,6].Value.ToString()+"\n");

                        }else{

                            Console.WriteLine(int.Parse(ex.Cells[cont,6].Value.ToString())*int.Parse(ex.Cells[cont,5].Value.ToString())+"\t");

                        }

                    }

                    cont++;

                }while(ex.Cells[cont,1].Value!=null);







            do{

                    Console.Write("\nDeseja realizar um novo cadastro? (S ou N)");

                    op1 = Console.ReadLine();

                } while (op1!="S" && op1!="N" && op1!="s" && op1!="n");

            } while(op1=="S" || op1=="s");

        }

        public void RealizarVendas(){
            int cont=1;
            string placa;
            string doc;
            int posPlaca;
            int parcelas=0;
            int linha;
            double valor = 0;

            ExibeProdutosNaoVendidos();

            ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Carro.xls");

            do{
                do{
                    Console.Write("\n\nInforme placa do carro a ser comprado: ");    
                    placa = Console.ReadLine();
                    Console.WriteLine(placa);
                }while(!rgx.IsMatch(placa));

                posPlaca = ValidarCompra(placa);
                Console.WriteLine(posPlaca);
            }while(posPlaca==0);

            cont=posPlaca;

            do{            
            Console.WriteLine("Forma de pagamento: (VISTA / PRAZO) ");
            formaPagamento = Console.ReadLine();
            }while(formaPagamento.ToUpper()!="VISTA" && formaPagamento.ToUpper()!="PRAZO");

            if(formaPagamento.ToUpper()=="VISTA"){
                Console.WriteLine(ex.Cells[cont,6].Value.ToString());
                valor = double.Parse(ex.Cells[cont,6].Value.ToString())-((double.Parse(ex.Cells[cont,6].Value.ToString())*5)/100);
            }
            else
            {
                Console.Write("Quantidade de Parcelas: ");
                parcelas = int.Parse(Console.ReadLine());
                valor = double.Parse(ex.Cells[cont,6].Value.ToString())/parcelas;
            }
            
            do{
                Console.WriteLine("Informe documento do cliente: ");           
                doc = Console.ReadLine();
                linha = cliente.PesquisaDocumento(doc); 

                if(linha==0){
                    Console.WriteLine("Cliente não cadastrado!");  
                    cliente.CadastrarClientes();
                }
            }while(linha==0);

            ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Carro.xls");

            Console.WriteLine(cont);
            ex.Cells[cont,7].Value = "1";
            ex.ActiveWorkbook.Save();
            ex.Quit();
            
            //CADASTRO DA VENDA

            ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Venda.xls");

            ex.Cells[linha,1].Value = doc;
            ex.Cells[linha,2].Value = placa;
            ex.Cells[linha,3].Value = DateTime.Now;
            ex.Cells[linha,4].Value = formaPagamento;
            ex.Cells[linha,5].Value = parcelas;
            ex.Cells[linha,6].Value = valor;

            ex.ActiveWorkbook.Save();
            ex.Quit();

            Console.WriteLine("Venda realizada!");
        }

        public void ExibeProdutosNaoVendidos(){
            ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Carro.xls");
            int cont=1;

            do{
                    if(ex.Cells[cont,7].Value.ToString() == "0"){
                        
                        Console.Write(ex.Cells[cont,1].Value+"\t");
                        Console.Write(ex.Cells[cont,2].Value+"-");
                        Console.Write(ex.Cells[cont,3].Value+"\t");
                        Console.Write(ex.Cells[cont,4].Value+@"\");
                        Console.Write(ex.Cells[cont,5].Value+"\t");
                        Console.Write("R$"+ex.Cells[cont,6].Value+"\n");

                    }
                cont++;
            }while(ex.Cells[cont,1].Value!=null);

            ex.Quit();
        }

        public int ValidarCompra(string placa){
            ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Carro.xls");
            int cont=1;

            do{
                if(ex.Cells[cont,1].Value.ToString() == placa){        
                    if(ex.Cells[cont,7].Value.ToString()=="1"){
                        Console.WriteLine("Carro já foi comprado!");
                        return 0;
                    }
                    else
                    {
                        return cont;
                    }
                }
                cont++;
            }while(ex.Cells[cont,1].Value!=null);

            //Voltando para posição da placa
            Console.WriteLine("Carro não encontrado!");
            return 0;
        }
    }
}