using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using namasdev.Core.Exceptions;
using namasdev.Core.Types;
using namasdev.Core.Validation;

namespace namasdev.Excel
{
    public class HojaExcel
    {
        public HojaExcel(ExcelWorkbook workbook, string nombre)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException("workbook");
            }
            if (String.IsNullOrWhiteSpace(nombre))
            {
                throw new ArgumentNullException("nombre");
            }

            Workbook = workbook;
            Nombre = nombre;

            Hoja = Workbook.Worksheets[nombre];
            if (Hoja == null)
            {
                throw new ExcepcionMensajeAlUsuario($"La hoja '{nombre}' no existe.");
            }
        }

        public HojaExcel(ExcelWorkbook workbook, int numero = 1)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException("workbook");
            }

            Workbook = workbook;

            Hoja = Workbook.Worksheets[numero];
            if (Hoja == null)
            {
                throw new ExcepcionMensajeAlUsuario($"La hoja {numero} no existe.");
            }

            Nombre = Hoja.Name;
        }

        public ExcelWorkbook Workbook { get; private set; }
        public string Nombre { get; private set; }
        public ExcelWorksheet Hoja { get; private set; }

        /// <summary>
        /// Valida que la hoja tenga los encabezados especificados en la fila especificada (por default, busca en la fila 1).
        /// </summary>
        /// <param name="encabezados"></param>
        /// <param name="columna"></param>
        /// <param name="fila"></param>
        public void ValidarEncabezados(string[] encabezados, int columna = 1, int fila = 1)
        {
            var errores = new List<string>();

            int cantidadEncabezados = encabezados.Length;
            string encabezadoBuscado;
            ExcelRange celda;
            for (int i = 0; i < cantidadEncabezados; i++)
            {
                encabezadoBuscado = encabezados[i];
                celda = Hoja.Cells[fila, i + columna];

                if (!String.Equals(celda.Text.Trim().ToLower(), encabezadoBuscado.Trim().ToLower()))
                {
                    errores.Add($"{encabezadoBuscado} ({celda.Address})");
                }
            }

            if (errores.Any())
            {
                throw new ExcepcionMensajeAlUsuario($"[{Nombre}] Encabezados no encontrados: {String.Join(", ", errores)}.");
            }
        }

        public void AplicarEstilo(int fila, int columna,
            ExcelHorizontalAlignment? alineacionHorizontal = null, ExcelVerticalAlignment? alineacionVertical = null,
            bool? textoEnNegrita = null, bool? autoAjustar = null, Color? colorTexto = null, Color? colorFondo = null, Color? colorBorde = null)
        {
            var rango = Hoja.Cells[fila, columna];
            AplicarEstilo(rango, alineacionHorizontal, alineacionVertical, textoEnNegrita, autoAjustar, colorTexto, colorFondo, colorBorde);
        }

        public void AplicarEstilo(int filaDesde, int columnaDesde, int filaHasta, int columnaHasta,
            ExcelHorizontalAlignment? alineacionHorizontal = null, ExcelVerticalAlignment? alineacionVertical = null,
            bool? textoEnNegrita = null, bool? autoAjustar = null, Color? colorTexto = null, Color? colorFondo = null, Color? colorBorde = null)
        {
            var rango = Hoja.Cells[filaDesde, columnaDesde, filaHasta, columnaHasta];
            AplicarEstilo(rango, alineacionHorizontal, alineacionVertical, textoEnNegrita, autoAjustar, colorTexto, colorFondo, colorBorde);
        }

        public void AplicarEstilo(string rangoCeldas,
            ExcelHorizontalAlignment? alineacionHorizontal = null, ExcelVerticalAlignment? alineacionVertical = null,
            bool? textoEnNegrita = null, bool? autoAjustar = null, Color? colorTexto = null, Color? colorFondo = null, Color? colorBorde = null)
        {
            AplicarEstilo(rangoCeldas, alineacionHorizontal, alineacionVertical, textoEnNegrita, autoAjustar, colorTexto, colorFondo, colorBorde);
        }

        public void AplicarEstilo(ExcelRange rango,
            ExcelHorizontalAlignment? alineacionHorizontal = null, ExcelVerticalAlignment? alineacionVertical = null,
            bool? textoEnNegrita = null, bool? autoAjustar = null, Color? colorTexto = null, Color? colorFondo = null, Color? colorBorde = null)
        {
            if (rango == null)
                throw new ArgumentNullException("rango");

            if (colorTexto.HasValue)
            {
                rango.Style.Font.Color.SetColor(colorTexto.Value);
            }

            if (colorFondo.HasValue)
            {
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Style.Fill.BackgroundColor.SetColor(colorFondo.Value);
            }

            if (colorBorde.HasValue)
            {
                var border = rango.Style.Border;
                border.Top.Style = border.Bottom.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
                border.Top.Color.SetColor(colorBorde.Value);
                border.Bottom.Color.SetColor(colorBorde.Value);
                border.Left.Color.SetColor(colorBorde.Value);
                border.Right.Color.SetColor(colorBorde.Value);
            }

            if (textoEnNegrita.HasValue)
            {
                rango.Style.Font.Bold = textoEnNegrita.Value;
            }

            if (alineacionHorizontal.HasValue)
            {
                rango.Style.HorizontalAlignment = alineacionHorizontal.Value;
            }

            if (alineacionVertical.HasValue)
            {
                rango.Style.VerticalAlignment = alineacionVertical.Value;
            }

            if (autoAjustar == true)
            {
                rango.AutoFitColumns();
                rango.Style.WrapText = true;
            }
        }
    }

    public abstract class RegistroExcel
    {
        private List<string> _errores;

        public RegistroExcel(ExcelWorksheet hoja, int fila,
            bool incluirNombreHojaEnError = true)
        {
            if (hoja == null)
            {
                throw new ArgumentNullException("hoja");
            }

            Hoja = hoja;
            Fila = fila;

            _errores = new List<string>();

            IncluirNombreHojaEnError = incluirNombreHojaEnError;
        }

        public ExcelWorksheet Hoja { get; private set; }
        public int Fila { get; private set; }
        public IEnumerable<string> Errores
        {
            get { return _errores.AsReadOnly(); }
        }

        public bool IncluirNombreHojaEnError { get; set; }

        public virtual bool EsRegistroVacio { get; }

        public bool EsValido
        {
            get { return _errores.Count == 0; }
        }

        protected string ObtenerString(int columna, string descripcionDato, 
            bool esRequerido = true,
            int? tamañoMaximo = null, int? tamañoExacto = null,
            bool aplicarTrim = false)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Text.ValueNotEmptyOrNull();

            if (!Validador.ValidarString(valor, descripcionDato, esRequerido,
                out string mensajeError,
                tamañoMaximo: tamañoMaximo, tamañoExacto: tamañoExacto))
            {
                AgregarError(celda, mensajeError);
            }
            else if (aplicarTrim)
            {
                valor = valor.Trim();
            }

            return valor;
        }

        protected int? ObtenerInt(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Value;

            if (valor != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(valor)))
            {
                try
                {
                    return Convert.ToInt32(valor);
                }
                catch (Exception)
                {
                    AgregarError(celda, String.Format(Validador.TIPO_NO_VALIDO_TEXTO_FORMATO, descripcionDato, "Número entero"));
                    return null;
                }
            }
            else
            {
                if (esRequerido)
                {
                    AgregarError(celda, String.Format(Validador.REQUERIDO_TEXTO_FORMATO, descripcionDato));
                }

                return null;
            }
        }

        protected string ObtenerStringYValidarFormatoCorreo(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            string strValor = ObtenerString(columna, descripcionDato, esRequerido);
            if (!string.IsNullOrEmpty(strValor))
            {
                string mensajeError;
                if (!Validador.ValidarEmail(strValor, "Correo electrónico", out mensajeError))
                {
                    AgregarError(celda, mensajeError);
                }
            }

            return strValor;
        }

        protected short? ObtenerShort(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Value;

            if (valor != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(valor)))
            {
                try
                {
                    return Convert.ToInt16(valor);
                }
                catch (Exception)
                {
                    AgregarError(celda, String.Format(Validador.TIPO_NO_VALIDO_TEXTO_FORMATO, descripcionDato, "Número entero corto"));
                    return null;
                }
            }
            else
            {
                if (esRequerido)
                    AgregarError(celda, String.Format(Validador.REQUERIDO_TEXTO_FORMATO, descripcionDato));

                return null;
            }
        }

        protected long? ObtenerLong(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Value;

            if (valor != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(valor)))
            {
                try
                {
                    return Convert.ToInt64(valor);
                }
                catch (Exception)
                {
                    AgregarError(celda, String.Format(Validador.TIPO_NO_VALIDO_TEXTO_FORMATO, descripcionDato, "Número entero largo"));
                    return null;
                }
            }
            else
            {
                if (esRequerido)
                    AgregarError(celda, String.Format(Validador.REQUERIDO_TEXTO_FORMATO, descripcionDato));

                return null;
            }
        }

        protected double? ObtenerDouble(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Value;

            if (valor != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(valor)))
            {
                try
                {
                    return Convert.ToDouble(valor);
                }
                catch (Exception)
                {
                    AgregarError(celda, String.Format(Validador.TIPO_NO_VALIDO_TEXTO_FORMATO, descripcionDato, "Número"));
                    return null;
                }
            }
            else
            {
                if (esRequerido)
                    AgregarError(celda, String.Format(Validador.REQUERIDO_TEXTO_FORMATO, descripcionDato));

                return null;
            }
        }

        protected decimal? ObtenerDecimal(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Value;

            if (valor != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(valor)))
            {
                try
                {
                    return Convert.ToDecimal(valor);
                }
                catch (Exception)
                {
                    AgregarError(celda, String.Format(Validador.TIPO_NO_VALIDO_TEXTO_FORMATO, descripcionDato, "Número"));
                    return null;
                }
            }
            else
            {
                if (esRequerido)
                    AgregarError(celda, String.Format(Validador.REQUERIDO_TEXTO_FORMATO, descripcionDato));

                return null;
            }
        }

        protected DateTime? ObtenerDateTime(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Value;

            if (valor != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(valor).Replace("-", "")))
            {
                try
                {
                    return Convert.ToDateTime(valor);
                }
                catch (Exception)
                {
                    try
                    {
                        double dato = Convert.ToDouble(valor);
                        return DateTime.FromOADate(dato);
                    }
                    catch (Exception)
                    {
                        AgregarError(celda, String.Format(Validador.TIPO_NO_VALIDO_TEXTO_FORMATO, descripcionDato, "Fecha/Hora"));
                        return null;
                    }
                }
            }
            else
            {
                if (esRequerido)
                    AgregarError(celda, String.Format(Validador.REQUERIDO_TEXTO_FORMATO, descripcionDato));

                return null;
            }

        }

        protected TimeSpan? ObtenerTimeSpan(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Value;

            if (valor != null
                && !String.IsNullOrWhiteSpace(Convert.ToString(valor).Replace("-", "")))
            {
                try
                {
                    return TimeSpan.Parse(Convert.ToString(valor));
                }
                catch (Exception)
                {
                    try
                    {
                        return Convert.ToDateTime(valor).TimeOfDay;
                    }
                    catch (Exception)
                    {
                        AgregarError(celda, String.Format(Validador.TIPO_NO_VALIDO_TEXTO_FORMATO, descripcionDato, "Hora"));
                        return null;
                    }
                }
            }
            else
            {
                if (esRequerido)
                    AgregarError(celda, String.Format(Validador.REQUERIDO_TEXTO_FORMATO, descripcionDato));

                return null;
            }
        }

        protected bool? ObtenerBoolean(int columna, string descripcionDato, bool esRequerido = true)
        {
            var celda = ObtenerCelda(columna);
            var valor = celda.Value;

            if (valor != null)
            {
                bool? resultado = valor as bool?;
                if (resultado.HasValue)
                {
                    return resultado.Value;
                }

                var strValor = valor.ToString();
                return string.Equals(strValor, Formateador.TEXTO_SI_SIN_ACENTO, StringComparison.CurrentCultureIgnoreCase)
                    || string.Equals(strValor, Formateador.TEXTO_SI_CON_ACENTO, StringComparison.CurrentCultureIgnoreCase);
            }
            else
            {
                if (esRequerido)
                    AgregarError(celda, String.Format(Validador.REQUERIDO_TEXTO_FORMATO, descripcionDato));

                return null;
            }
        }

        protected short? ObtenerMesNumero(int columna, string descripcionDato, bool esRequerido = true)
        {
            var valor = ObtenerString(columna, descripcionDato, esRequerido: esRequerido);
            if (String.IsNullOrWhiteSpace(valor))
            {
                return null;
            }

            try
            {
                short mes;
                if (short.TryParse(valor, out mes))
                {
                    return mes;
                }

                var meses = new List<string>(CultureInfo.CurrentCulture.DateTimeFormat.MonthNames);
                int indice = meses.IndexOf(valor.ToLower());
                if (indice >= 0)
                {
                    return (short)(indice + 1);
                }
                else
                {
                    throw new Exception("Mes inexistente.");
                }
            }
            catch (Exception)
            {
                AgregarError(ObtenerCelda(columna), $"{descripcionDato} no es un mes válido.");
                return null;
            }
        }

        protected ExcelRange ObtenerCelda(int columna)
        {
            return Hoja.Cells[Fila, columna];
        }

        protected void AgregarError(ExcelRange celda, string error)
        {
            if (celda != null)
            {
                _errores.Add($"[{(IncluirNombreHojaEnError ? celda.FullAddress : celda.Address)}] {error}");
            }
            else
            {
                _errores.Add(error);
            }
        }

        /// <summary>
        /// Agrega el valor especificado a la celda (fila, columna) de la hoja.
        /// </summary>
        /// <param name="columna"></param>
        /// <param name="valor"></param>
        protected void AgregarValorACelda(int columna, object valor)
        {
            Hoja.Cells[Fila, columna].Value = valor;
        }

        protected void AgregarValorACeldaWrapped(int columna, object valor)
        {
            Hoja.Cells[Fila, columna].Value = valor;
            Hoja.Cells[Fila, columna].Style.WrapText = true;
        }

        protected void AgregarValorConFormatoMonedaACelda(int columna, object valor)
        {
            AgregarValorConFormatoNumericoACelda(columna, valor, "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-");
        }

        protected void AgregarValorConFormatoPorcentajeACelda(int columna, object valor)
        {
            AgregarValorConFormatoNumericoACelda(columna, valor, "0%");
        }

        protected void AgregarValorConFormatoFechaCortaACelda(int columna, object valor)
        {
            AgregarValorConFormatoNumericoACelda(columna, valor, DateTimeFormatInfo.CurrentInfo.ShortDatePattern);
        }

        protected void AgregarValorConFormatoFechaLargaACelda(int columna, object valor)
        {
            AgregarValorConFormatoNumericoACelda(columna, valor, "dddd, d \\d\\e MMMM \\d\\e yyyy");
        }

        protected void AgregarValorConFormatoNumericoACelda(int columna, object valor, string formato)
        {
            Hoja.Cells[Fila, columna].Value = valor;
            Hoja.Cells[Fila, columna].Style.Numberformat.Format = formato;
        }

        protected void AplicarEstilo(int columna,
            ExcelHorizontalAlignment? alineacionHorizontal = null, Color? colorTexto = null, Color? colorFondo = null, Color? colorBorde = null)
        {
            AplicarEstilo(columna, columna, alineacionHorizontal, colorTexto, colorFondo, colorBorde);
        }

        protected void AplicarEstilo(int columnaDesde, int columnaHasta,
            ExcelHorizontalAlignment? alineacionHorizontal = null, Color? colorTexto = null, Color? colorFondo = null, Color? colorBorde = null)
        {
            var rango = Hoja.Cells[Fila, columnaDesde, Fila, columnaHasta];

            if (colorTexto.HasValue)
                rango.Style.Font.Color.SetColor(colorTexto.Value);

            if (colorFondo.HasValue)
            {
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Style.Fill.BackgroundColor.SetColor(colorFondo.Value);
            }

            if (colorBorde.HasValue)
                rango.Style.Border.BorderAround(ExcelBorderStyle.Thin, colorBorde.Value);

            if (alineacionHorizontal.HasValue)
                rango.Style.HorizontalAlignment = alineacionHorizontal.Value;
        }
    }

    public class ExcelHelper
    {
        public static void ActualizarValoresDeNombreAdministrado<T>(ExcelWorkbook workbook, string nombreHoja, string nombreAdministrado, int columna, IEnumerable<T> valores,
            Func<T, object> mapeoValor = null, int filaInicio = 1, int filaMaxima = 9999)
        {
            var hoja = workbook.Worksheets[nombreHoja];
            hoja.Cells[filaInicio, columna, filaMaxima, columna].Value = null;

            int fila = filaInicio - 1;
            foreach (var valor in valores)
            {
                fila++;

                if (mapeoValor != null)
                {
                    hoja.Cells[fila, columna].Value = mapeoValor(valor);
                }
                else
                {
                    hoja.Cells[fila, columna].Value = valor;
                }
            }
            workbook.Names[nombreAdministrado].Address = hoja.Cells[filaInicio, columna, fila, columna].FullAddress;
        }
    }

    public class Encabezado
    {
        public Encabezado(int columna, string texto)
        {
            Columna = columna;
            Texto = texto;
        }

        public int Columna { get; private set; }
        public string Texto { get; private set; }
    }
}
