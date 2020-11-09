using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RM
{
    class Variavel
    {
        private int nRm;
        private int nRmf;
        private string nNome;
        private string nRg;
        private string nSp;
        private string nOp;
        private string nCam;

        public int Matricula
        {
            get { return nRm; }
            set { nRm = value; }
        }

        public int MatriculaFinal
        {
            get { return nRmf; }
            set { nRmf = value; }
        }

        public string Nome
        {
            get { return nNome; }
            set { nNome = value; }
        }

        public string Registro
        {
            get { return nRg; }
            set { nRg = value; }
        }

        public string SpDoc
        {
            get { return nSp; }
            set { nSp = value; }
        }

        public string Opcao
        {
            get { return nOp; }
            set { nOp = value; }
        }

        public string Caminho
        {
            get { return nCam; }
            set { nCam = value; }
        }
    }
}
