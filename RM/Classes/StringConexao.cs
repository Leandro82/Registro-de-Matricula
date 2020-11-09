using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RM
{
    class StringConexao
    {
        string caminho;

        public string Endereco()
        {
            caminho = "Persist Security Info=false;SERVER=10.66.121.42;DATABASE=agenda;UID=secac;pwd=secac;Allow User Variables=True;Convert Zero Datetime=True;default command timeout=0";
            return caminho;
        }
    }
}
