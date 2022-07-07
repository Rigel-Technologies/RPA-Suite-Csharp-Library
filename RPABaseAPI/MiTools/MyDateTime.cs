using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MiTools
{
    public static class MyDateTime : Object
    {
        private enum ModoTruncar { mtTzquierda, mTDerecha, mTFraccion };
        private static readonly char[] CaracteresEspecial = { 'Y', 'M', 'D', 'H', 'N', 'S', 'F' };
        private static readonly char[] CaracteresDigitos = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
        private static readonly char[] CaracteresSeparador = { '/', '-', '.' };

        private static string IgualesM(char car, ref int i, string cadena)
        {
            string resultado = string.Empty;

            while ((i < cadena.Length) && (car == cadena.ElementAt<char>(i)))
            {
                resultado = resultado + cadena.ElementAt<char>(i);
                i++;
            }
            return resultado;
        }
        private static int LeerNumeroM(int maximo, ref int i, string cadena)
        {
            string n = string.Empty;

            while ((i < cadena.Length) && (maximo > n.Length) && cadena.ElementAt<char>(i).IsIn(CaracteresDigitos))
            {
                n = n + cadena.ElementAt<char>(i);
                i++;
            }
            return Int32.Parse(n);
        }
        private static int SacarCampoM(ref int i, ref int j, string formato, string valor, out int longitud)
        {
            int resultado;
            string subformato;

            subformato = IgualesM(formato.ElementAt<char>(i), ref i, formato);
            longitud = subformato.Length;
            if (i < formato.Length)
            {
                if (formato.ElementAt<char>(i).IsIn(CaracteresEspecial))
                    resultado = LeerNumeroM(longitud, ref j, valor);
                else resultado = LeerNumeroM(100, ref j, valor);
            }
            else resultado = LeerNumeroM(100, ref j, valor);
            return resultado;
        }
        private static string SacarCampoM(ref int i, string formato, int precision, ModoTruncar truncar, int valor)
        {
            string resultado = string.Empty, subformato, lsValor;
            int longitud, liTrunc;

            subformato = IgualesM(formato.ElementAt<char>(i), ref i, formato);
            longitud = subformato.Length;
            lsValor = MyObject.AlignRight(valor.ToString(), '0', precision);
            if (lsValor.Length == longitud) resultado = lsValor;
            else if (lsValor.Length < longitud) resultado = MyObject.AlignRight(lsValor, '0', longitud);
            else if (truncar == ModoTruncar.mTDerecha)
            {
                liTrunc = longitud.Maximum(2);
                resultado = lsValor.Substring(lsValor.Length - liTrunc - 1, liTrunc);
            }
            else if (truncar == ModoTruncar.mtTzquierda)
            {
                liTrunc = longitud.Maximum(precision);
                resultado = lsValor.Substring(0, liTrunc);
            }
            else resultado = lsValor.Substring(0, longitud.Minimum(lsValor.Length));
            return resultado;
        }

        /* The format used is this:
             - yyyy for 4-digit years, yy for 2-digit years
             - MM for 2-digit months, m for 1-digit
             - d y dd
             - h y hh
             - n y nn for minutes
             - s y ss for seconds
             - f y ff y fff for miliseconds
         *****************************************************************/
        public static string FormatDateTime(string format, DateTime value)
        {
            string resultado = string.Empty;
            string lsUFormato;
            int i;

            try
            {
                lsUFormato = format.ToUpper();
                i = 0;
                while (i < lsUFormato.Length)
                {
                    switch (lsUFormato.ElementAt<char>(i))
                    {
                        case 'Y':
                            resultado = resultado + SacarCampoM(ref i, lsUFormato, 4, ModoTruncar.mTDerecha, value.Year);
                            break;
                        case 'M':
                            resultado = resultado + SacarCampoM(ref i, lsUFormato, 2, ModoTruncar.mtTzquierda, value.Month);
                            break;
                        case 'D':
                            resultado = resultado + SacarCampoM(ref i, lsUFormato, 2, ModoTruncar.mtTzquierda, value.Day);
                            break;
                        case 'H':
                            resultado = resultado + SacarCampoM(ref i, lsUFormato, 2, ModoTruncar.mtTzquierda, value.Hour);
                            break;
                        case 'N':
                            resultado = resultado + SacarCampoM(ref i, lsUFormato, 2, ModoTruncar.mtTzquierda, value.Minute);
                            break;
                        case 'S':
                            resultado = resultado + SacarCampoM(ref i, lsUFormato, 2, ModoTruncar.mtTzquierda, value.Second);
                            break;
                        case 'F':
                            resultado = resultado + SacarCampoM(ref i, lsUFormato, 3, ModoTruncar.mTFraccion, value.Millisecond);
                            break;
                        default:
                            resultado = resultado + format.ElementAt<char>(i);
                            i++;
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                MyObject.Coroner.write("MyDateTime.FormatDateTime(string, DateTime)", e);
                throw;
            }
            return resultado;
        }
        public static DateTime SetFromFormat(string format, string value)
        {
            DateTime resultado;
            int ano = 0, mes = 0, dia = 0, hora = 0, minuto = 0, segundo = 0, mili = 0;
            int lsiglo, lano;
            bool lbOano = false, lbOmes = false, lbOdia = false;
            int lbOminuto = 0;
            int i, j, z, subformato;

            try
            {
                if (value == null) value = string.Empty;
                format = format.ToUpper();
                i = 0;
                while (i < format.Length)
                {
                    switch (format.ElementAt<char>(i))
                    {
                        case 'Y':
                            lbOano = true;
                            break;
                        case 'M':
                            lbOmes = true;
                            break;
                        case 'D':
                            lbOdia = true;
                            break;
                        case 'N':
                            lbOminuto = 1;
                            break;
                    }
                    i++;
                }
                i = 0;
                j = 0;
                while ((i < format.Length) && (j < value.Length))
                {
                    switch (format.ElementAt<char>(i))
                    {
                        case 'Y':
                            lano = DateTime.Now.Year;
                            lsiglo = lano / 100;
                            ano = SacarCampoM(ref i, ref j, format, value, out subformato);
                            lbOano = false;
                            if ((subformato == 2) && (ano < 100))
                                ano = lsiglo * 100 + ano;
                            break;
                        case 'M':
                            mes = SacarCampoM(ref i, ref j, format, value, out subformato);
                            lbOmes = false;
                            break;
                        case 'D':
                            dia = SacarCampoM(ref i, ref j, format, value, out subformato);
                            lbOdia = false;
                            break;
                        case 'H':
                            hora = SacarCampoM(ref i, ref j, format, value, out subformato);
                            lbOminuto++;
                            break;
                        case 'N':
                            minuto = SacarCampoM(ref i, ref j, format, value, out subformato);
                            lbOminuto = -1;
                            break;
                        case 'S':
                            segundo = SacarCampoM(ref i, ref j, format, value, out subformato);
                            break;
                        case 'F':
                            z = j;
                            mili = SacarCampoM(ref i, ref j, format, value, out subformato);
                            if (subformato > 3) throw new Exception("Format exceeds precision in milliseconds.");
                            if ((j - z) < 3)
                            {
                                mili = mili * (int)Math.Pow(10, 3 - (j - z));
                            }
                            break;
                        default:
                            if ((format.ElementAt<char>(i) == value.ElementAt<char>(j)) ||
                                (format.ElementAt<char>(i).IsIn(CaracteresSeparador) && value.ElementAt<char>(j).IsIn(CaracteresSeparador)))
                            {
                                i++;
                                j++;
                            }
                            else throw new Exception("The value '" + value + "' does not correspond to the specified format : '" + format + "'");
                            break;
                    }
                }
                if (lbOano || lbOmes || lbOdia || (lbOminuto == 2))
                    throw new Exception("The value '" + value + "' does not correspond to the specified format : '" + format + "'");
                resultado = new DateTime(ano, mes, dia, hora, minuto, segundo, mili);
            }
            catch (Exception e)
            {
                MyObject.Coroner.write("MyDateTime.SetFromFormat(string, string)", e);
                throw;
            }
            return resultado;
        }
        public static bool TryParse(string format, string value, out DateTime date)
        {
            bool result = false;

            try
            {
                date = SetFromFormat(format, value);
                result = true;
            }
            catch
            {
                date = DateTime.Today;
            }
            return result;
        }
        public static string FormatDateTime(this MyObject a, string formato, DateTime value)
        {
            return MyDateTime.FormatDateTime(formato, value);
        }
        public static DateTime ToDateTime(this MyObject a, string formato, string valor)
        {
            return MyDateTime.SetFromFormat(formato, valor);
        }
        public static DateTime Minimum(this DateTime a, DateTime b)
        {
            if (a < b) return a;
            else return b;
        }
        public static DateTime Maximum(this DateTime a, DateTime b)
        {
            if (a < b) return b;
            else return a;
        }
    }

    public static class MyInt
    {
        public static int Maximum(this int a, int b)
        {
            if (a > b) return a;
            else return b;
        }
        public static int Minimum(this int a, int b)
        {
            if (a > b) return b;
            else return a;
        }
    }
    public static class MyDouble
    {
        public static double Maximum(this double a, double b)
        {
            if (a > b) return a;
            else return b;
        }
        public static double Minimum(this double a, double b)
        {
            if (a > b) return b;
            else return a;
        }
    }
}
