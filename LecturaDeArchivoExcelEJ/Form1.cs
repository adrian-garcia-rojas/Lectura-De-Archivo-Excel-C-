using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


//No olvides agregar al proyecto la referencia Microsoft.Office.Interop.Excel e incluirla en el proyecto

namespace LecturaDeArchivoExcelEJ
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void botonBuscar_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Buscar";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "xls";
            openFileDialog1.Filter = "Archivos de Excel (*.xls)|*.xls|Todos (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void botonLeerExcel_Click(object sender, EventArgs e)
        {
            String rutaDelArchivo = textBox1.Text;
            if (rutaDelArchivo.Length > 0)
            {
                recorrerArchivoExcel(rutaDelArchivo);
            }
            else
            {
                MessageBox.Show("Primero debes seleccionar un archivo de Excel");
            }
        }

        public void recorrerArchivoExcel(String rutaDelArchivo)
        {
            Excel.Application excelApplicacion;
            Excel.Workbook excelWorkBook;
            Excel.Worksheet excelHoja;
            Excel.Range excelRango;

            try
            {
                excelApplicacion = new Excel.Application();
                excelWorkBook = excelApplicacion.Workbooks.Open(rutaDelArchivo, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                excelHoja = (Excel.Worksheet)excelWorkBook.Worksheets.get_Item(1);

                excelRango = excelHoja.UsedRange;
                Int32 rangoEnColumnas = excelRango.Columns.Count;
                Int32 rangoEnFilas = excelRango.Rows.Count;
                String valorDeCelda = "";
                String cadena = "";

                label4.Text = "Filas: " + rangoEnFilas + ", Columnas: " + rangoEnColumnas;

                if (rangoEnFilas > 0 && rangoEnColumnas > 0)
                {                 
                    for (int posicionFila = 1; posicionFila <= rangoEnFilas; posicionFila++)
                    {
                        for (int posicionColumna = 1; posicionColumna <= rangoEnColumnas; posicionColumna++)
                        {
                            valorDeCelda = (string)(Convert.ToString(((excelRango.Cells[posicionFila, posicionColumna] as Excel.Range).Value2)));                            
                            label2.Text = "Dato: " + valorDeCelda;
                            cadena = cadena + valorDeCelda + " ";
                            if (posicionColumna == 3)
                            {
                                richTextBox1.Text += cadena + Environment.NewLine;
                                cadena = "";
                            }
                        }

                        label3.Text = "Filas: " + posicionFila + " de " + rangoEnFilas;
                    }
                    MessageBox.Show("Recorrido finalizado");
                }
                else
                {
                    MessageBox.Show("El archivo no contiene información");
                }

                excelWorkBook.Close(true, null, null);
                excelApplicacion.Quit();

                releaseObject(excelHoja);
                releaseObject(excelWorkBook);
                releaseObject(excelApplicacion);
            }
            catch (Exception e)
            {
                MessageBox.Show("Debes contar con Microsoft Office en tu equipo!" + Environment.NewLine + e);                
            }

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
                MessageBox.Show("No es posible liberar objeto" + e);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
