using System;
using NetOffice.ExcelApi;

namespace concessionaria_classes
{
    public class Validacao
    {

            public int ValidarCPF(String docCliente)
            {
            int[] chave1 = {10,9,8,7,6,5,4,3,2};
            int[] chave2 = {11,10,9,8,7,6,5,4,3,2};

            string primeiroDigito, segundoDigito;

            if(docCliente=="00000000000" || docCliente=="11111111111" || docCliente=="22222222222"
            || docCliente=="33333333333" || docCliente=="44444444444" || docCliente=="55555555555"
            || docCliente=="66666666666" || docCliente=="77777777777" || docCliente=="88888888888"
            || docCliente=="99999999999")
            {
                Console.WriteLine("CPF inválido!");
                return 0;
            }

            primeiroDigito = ValidaDigito(docCliente,chave1,1);

            if (primeiroDigito != docCliente.Substring(chave1.Length,1))
            {
                Console.WriteLine("CPF inválido!");
            }
            else
            {
                segundoDigito = ValidaDigito(docCliente,chave2,1);

            if (docCliente.EndsWith(segundoDigito) == true)
            {
                return 1;
            }
            else
            {
                Console.WriteLine("CPF inválido!");
            }
            }
        return 0;
        }

        public int ValidarCNPJ(String docCliente)
            {
            int[] chave1 = {5,4,3,2,9,8,7,6,5,4,3,2};
            int[] chave2 = {6,5,4,3,2,9,8,7,6,5,4,3,2};
            string primeiroDigito, segundoDigito;

            primeiroDigito = ValidaDigito(docCliente,chave1,2);

            if (primeiroDigito != docCliente.Substring(chave1.Length,1))
            {
                Console.WriteLine("CNPJ inválido!");
            }
            else
            {
                segundoDigito = ValidaDigito(docCliente,chave2,2);

                if (docCliente.EndsWith(segundoDigito) == true)
                {
                    return 1;
                }
                else
                {
                    Console.WriteLine("CNPJ inválido!");
                }
            }
        return 0;
        }
        private static string ValidaDigito(string doc, int[] chave, int tipoDoc)
        {
       int soma = 0, resto = 0;
       string tempdoc;
       tempdoc = doc.Substring(0,chave.Length);

       for(int i=0;i<chave.Length;i++){
                soma += Convert.ToInt16(tempdoc[i].ToString())*chave[i];
        }
            resto = soma % 11;

            if(resto<2)
            {
                return "0";        
            }
            else
            {
                return (11-resto).ToString();
            }

    }
    }
}