using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Planejamento_RPA
{
    class Program
    {
        static void Main(string[] args)
        {

            string diretorio = @"\\10.0.0.12\cred$\Dados\Planejamento\RPA";

            Processos proc = new Processos();

            proc.Acompanhamento_BV(diretorio);
            proc.WhatsApp_Resumo_BV(diretorio);


        }
    }
}
