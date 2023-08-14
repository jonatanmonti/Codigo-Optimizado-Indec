using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codigo_Optimizado_Indec
{
    public class ObraSieteItem : ObraCincoItem
    {

        private float hormigon; //Item Hormigon

        public float Hormigon
        {
            get { return hormigon; }
            set { hormigon = value; }
        }

        private float acero; //Item Acero

        public float Acero
        {
            get { return acero; }
            set { acero = value; }
        }

    }
}