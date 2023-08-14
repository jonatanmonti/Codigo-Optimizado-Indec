using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codigo_Optimizado_Indec
{
    public class ObraCincoItem
    {

        private float manoDeObra; //item Mano de Obra

        public float ManoDeObra
        {
            get { return manoDeObra; }
            set { manoDeObra = value; }
        }

        private float equipo; //item Equipo

        public float Equipo
        {
            get { return equipo; }
            set { equipo = value; }
        }

        private float asfaltosCombustibleLubricante; //Item Astaltos, combustibles y lubricantes

        public float AsfaltoCombustibleLubricante
        {
            get { return asfaltosCombustibleLubricante; }
            set { asfaltosCombustibleLubricante = value; }
        }

        private float gastoGeneral; //Item Gasto general

        public float GastoGeneral
        {
            get { return gastoGeneral; }
            set { gastoGeneral = value; }
        }

        private float costoFinanciero; //Item Costo Financiero

        public float CostoFinanciero
        {
            get { return costoFinanciero; }
            set { costoFinanciero = value; }
        }

        private float total; //Item Total

        public float Total
        {
            get { return total; }
            set { total = value; }
        }

    }
}