using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Net.WebRequestMethods;
using System.Runtime.InteropServices.ComTypes;

namespace AUTOCAD_Redesign
{
    public partial class Form1 : Form
    {
        //string explode = "(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")(command\"_explode\"\"all\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"layer\"\"0\"\"\")(command\"_change\"\"all\"\"\"\"properties\"\"Color\"\"Truecolor\"\"255,255,255\"\"\")";
        string explode = "(while (setq sset (ssget \"X\" \'((0 . \"INSERT\")))) (sssetfirst nil sset) (C:Burst)) (princ) (command\"_-purge\"\"B\"\"*\"\"N\")";

        public class formation_type
        {
            public const int type_2x2_5 = 1;
            public const int type_4x2_5 = 2;
            public const int type_4x4 = 3;
        }

        public class type_destination
        {
            public const int type_simple = 1;
            public const int type_double = 2;
        }

        public Form1()
        {
            InitializeComponent();
            InitializeOpenFileDialog();
            saveRoute.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";

            reporte.Text = "Estado: Inicialización del programa";

                //"C:\\Users\\JZ4874\\Desktop\\";
        }

        private void InitializeOpenFileDialog()
        {
            // Set the file dialog to filter for graphics files.
            this.openFileDialog1.Filter = "Excel |*.xlsx";
                //"All files (*.*)|*.*";

            //  Allow the user to select multiple images.
            this.openFileDialog1.Multiselect = true;
            //                   ^  ^  ^  ^  ^  ^  ^
            this.openFileDialog1.FileName = "";

            this.openFileDialog1.Title = "Buscar Archivos Excel";
        }

        private string convertRangeToString(Excel.Range inputRange)
        {
            return ((inputRange as Excel.Range).Value) == null ? "" : (string)((inputRange as Excel.Range).Value).ToString();
        }

        private void change_formacion(int type, Excel.Worksheet sheet, int i, int type_dest)
        {
            int col_function = 17;
            int row_init = 17;
            int row_fin = 110;

            if(type_dest == type_destination.type_double)
            {
                col_function = 26;
            }

            switch(type)
            {
                case formation_type.type_2x2_5:
                    for (int j = row_init; j < row_fin; j++)
                    {
                        sheet.Application.Goto(((Excel.Range)sheet.Cells[j, i]), true);
                        
                        if ((convertRangeToString(sheet.Cells[j, 10]) == "CONTINÚA EN HOJA N°:") || (convertRangeToString(sheet.Cells[j, 19]) == "CONTINÚA EN HOJA N°:"))
                        {
                            break;
                        }

                        if (convertRangeToString(sheet.Cells[j, i]) == "1")
                        {
                            reporte.Text = "Estado: Reemplazando formación 2x2.5 vena 1 por Marrón en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                            sheet.Cells[j, i] = "M";
                        }
                        else if (convertRangeToString(sheet.Cells[j, i]) == "2")
                        {
                            reporte.Text = "Estado: Reemplazando formación 2x2.5 vena 2 por Celeste en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                            sheet.Cells[j, i] = "C";
                        }
                    }
                    break;
                case formation_type.type_4x2_5:
                    for (int j = row_init; j < row_fin; j++)
                    {
                        sheet.Application.Goto(((Excel.Range)sheet.Cells[j, i]), true);

                        if ((convertRangeToString(sheet.Cells[j, 10]) == "CONTINÚA EN HOJA N°:") || (convertRangeToString(sheet.Cells[j, 19]) == "CONTINÚA EN HOJA N°:"))
                        {
                            break;
                        }

                        //((Excel.Range) sheet.Cells[j, i]).Select();
                        if (convertRangeToString(sheet.Cells[j,i]) != "")
                        {
                            //((Excel.Range)sheet.Cells[j, i]).Select();
                            if (convertRangeToString(sheet.Cells[j, col_function]).Contains("Fase R"))
                            {
                                DialogResult res;
                                res = MessageBox.Show("Desea cambiar el valor de la formación 4x2.5 de la fase R en la hoja " + sheet.Name + "?",
                                    "Reemplazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (res == DialogResult.Yes)
                                {
                                    sheet.Cells[j, i] = "C";
                                    reporte.Text = "Estado: Reemplazando formación 4x2.5 Fase R por Celeste en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                                }
                                else if (res == DialogResult.Cancel)
                                {
                                    System.Windows.Forms.Application.Exit();
                                }
                            }
                            else if (convertRangeToString(sheet.Cells[j, col_function]).Contains("Fase S"))
                            {
                                DialogResult res;
                                res = MessageBox.Show("Desea cambiar el valor de la formación 4x2.5 de la fase S en la hoja " + sheet.Name + "?",
                                    "Reemplazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (res == DialogResult.Yes)
                                {
                                    sheet.Cells[j, i] = "M";
                                    reporte.Text = "Estado: Reemplazando formación 4x2.5 Fase S por Marrón en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                                }
                                else if (res == DialogResult.Cancel)
                                {
                                    System.Windows.Forms.Application.Exit();
                                }
                            }
                            else if (convertRangeToString(sheet.Cells[j, col_function]).Contains("Fase T"))
                            {
                                DialogResult res;
                                res = MessageBox.Show("Desea cambiar el valor de la formación 4x2.5 de la fase T en la hoja " + sheet.Name + "?",
                                    "Reemplazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (res == DialogResult.Yes)
                                {
                                    sheet.Cells[j, i] = "R";
                                    reporte.Text = "Estado: Reemplazando formación 4x2.5 Fase T por Rojo en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                                }
                                else if (res == DialogResult.Cancel)
                                {
                                    System.Windows.Forms.Application.Exit();
                                }
                            }
                            else if (convertRangeToString(sheet.Cells[j, col_function]).Contains("Neutro"))
                            {
                                DialogResult res;
                                res = MessageBox.Show("Desea cambiar el valor de la formación 4x2.5 del neutro en la hoja " + sheet.Name + "?",
                                    "Reemplazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (res == DialogResult.Yes)
                                {
                                    sheet.Cells[j, i] = "N";
                                    reporte.Text = "Estado: Reemplazando formación 4x2.5 Neutro por Negro en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                                }
                                else if (res == DialogResult.Cancel)
                                {
                                    System.Windows.Forms.Application.Exit();
                                }
                            }
                            else
                            {
                                sheet.Cells[j, i].Interior.Color = 0x00FFFF;
                            }
                        }

                    }
                    break;
                case formation_type.type_4x4:
                    for (int j = row_init; j < row_fin; j++)
                    {
                        sheet.Application.Goto(((Excel.Range)sheet.Cells[j, i]), true);

                        if ((convertRangeToString(sheet.Cells[j, 10]) == "CONTINÚA EN HOJA N°:") || (convertRangeToString(sheet.Cells[j, 19]) == "CONTINÚA EN HOJA N°:"))
                        {
                            break;
                        }

                        if (convertRangeToString(sheet.Cells[j, i]) != "")
                        {
                            //((Excel.Range)sheet.Cells[j, i]).Select();
                            if (convertRangeToString(sheet.Cells[j, col_function]).Contains("Fase R"))
                            {
                                //DialogResult res;
                                //res = MessageBox.Show("Desea cambiar el valor de la formación 4x4 de la fase R en la hoja " + sheet.Name + "?",
                                //    "Reemplazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                //if (res == DialogResult.Yes)
                                //{
                                sheet.Cells[j, i] = "C";
                                reporte.Text = "Estado: Reemplazando formación 4x4 Fase R por Celeste en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                                //}
                                //else if (res == DialogResult.Cancel)
                                //{
                                //    System.Windows.Forms.Application.Exit();
                                //}
                            }
                            else if (convertRangeToString(sheet.Cells[j, col_function]).Contains("Fase S"))
                            {
                                //DialogResult res;
                                //res = MessageBox.Show("Desea cambiar el valor de la formación 4x4 de la fase S en la hoja " + sheet.Name + "?",
                                //    "Reemplazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                //if (res == DialogResult.Yes)
                                //{
                                sheet.Cells[j, i] = "M";
                                reporte.Text = "Estado: Reemplazando formación 4x4 Fase S por Marrón en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                                //}
                                //else if (res == DialogResult.Cancel)
                                //{
                                //    System.Windows.Forms.Application.Exit();
                                //}
                            }
                            else if (convertRangeToString(sheet.Cells[j, col_function]).Contains("Fase T"))
                            {
                                //DialogResult res;
                                //res = MessageBox.Show("Desea cambiar el valor de la formación 4x4 de la fase T en la hoja " + sheet.Name + "?",
                                //    "Reemplazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                //if (res == DialogResult.Yes)
                                //{
                                sheet.Cells[j, i] = "R";
                                reporte.Text = "Estado: Reemplazando formación 4x4 Fase T por Rojo en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                                //}
                                //else if (res == DialogResult.Cancel)
                                //{
                                //    System.Windows.Forms.Application.Exit();
                                //}
                            }
                            else if (convertRangeToString(sheet.Cells[j, col_function]).Contains("Neutro"))
                            {
                                //DialogResult res;
                                //res = MessageBox.Show("Desea cambiar el valor de la formación 4x4 del neutro en la hoja " + sheet.Name + "?",
                                //    "Reemplazo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                //if (res == DialogResult.Yes)
                                //{
                                sheet.Cells[j, i] = "N";
                                reporte.Text = "Estado: Reemplazando formación 4x4 Neutro por Negro en fila " + j + ", col " + i + ", hoja " + sheet.Name;
                                //}
                                //else if (res == DialogResult.Cancel)
                                //{
                                //    System.Windows.Forms.Application.Exit();
                                //}
                            }
                            else
                            {
                                sheet.Cells[j, i].Interior.Color = 0x00FFFF;
                            }
                        }
                    }

                    break;
                default:
                    MessageBox.Show("La formación no es correcta", "Error");
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(System.IO.Directory.Exists(saveRoute.Text))
            {
                if (System.IO.File.Exists(nameFile.Text) && (Path.GetExtension(nameFile.Text) == ".xlsx"))
                {
                    reporte.Text = "Estado: Copiando el archivo a la ruta indicada";

                    System.IO.File.Copy(nameFile.Text, saveRoute.Text + Path.GetFileName(nameFile.Text), true);

                    reporte.Text = "Estado: Abriendo el documento Excel";

                    Excel.Application app = new Excel.Application();
                    app.Visible = true;
                    app.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                    Excel.Workbook book = app.Workbooks.Open(saveRoute.Text + Path.GetFileName(nameFile.Text));

                    foreach (Excel.Worksheet sheet in book.Worksheets)
                    {
                        if (sheet.Name != "CARATULA")
                        {
                            reporte.Text = "Estado: Procesando el archivo, hoja " + sheet.Name;

                            if (convertRangeToString(sheet.Cells[12, 21]) == "CABLE N°")
                            {
                                for (int i = 22; i < 39; i++)
                                {
                                    sheet.Application.Goto(((Excel.Range)sheet.Cells[3, i]), true);
                                    //((Excel.Range) sheet.Cells[3, i]).Select();
                                    if (convertRangeToString(sheet.Cells[3, i]) == "2x2.5")
                                    {
                                        change_formacion(formation_type.type_2x2_5, sheet, i, type_destination.type_simple);
                                    }
                                    else if (convertRangeToString(sheet.Cells[3, i]) == "4x4")
                                    {
                                        change_formacion(formation_type.type_4x4, sheet, i, type_destination.type_simple);
                                    }
                                    else if (convertRangeToString(sheet.Cells[3, i]) == "4x2.5")
                                    {
                                        change_formacion(formation_type.type_4x2_5, sheet, i, type_destination.type_simple);
                                    }
                                }
                            }
                            else if (convertRangeToString(sheet.Cells[12, 18]) == "CABLE N°")
                            {
                                for (int i = 10; i < 18; i++)
                                {
                                    if (convertRangeToString(sheet.Cells[3, i]) == "2x2.5")
                                    {
                                        change_formacion(formation_type.type_2x2_5, sheet, i, type_destination.type_double);
                                    }
                                    else if (convertRangeToString(sheet.Cells[3, i]) == "4x4")
                                    {
                                        change_formacion(formation_type.type_4x4, sheet, i, type_destination.type_double);
                                    }
                                    else if (convertRangeToString(sheet.Cells[3, i]) == "4x2.5")
                                    {
                                        change_formacion(formation_type.type_4x2_5, sheet, i, type_destination.type_double);
                                    }
                                }

                                for (int i = 30; i < 38; i++)
                                {
                                    if (convertRangeToString(sheet.Cells[3, i]) == "2x2.5")
                                    {
                                        change_formacion(formation_type.type_2x2_5, sheet, i, type_destination.type_double);
                                    }
                                    else if (convertRangeToString(sheet.Cells[3, i]) == "4x4")
                                    {
                                        change_formacion(formation_type.type_4x4, sheet, i, type_destination.type_double);
                                    }
                                    else if (convertRangeToString(sheet.Cells[3, i]) == "4x2.5")
                                    {
                                        change_formacion(formation_type.type_4x2_5, sheet, i, type_destination.type_double);
                                    }
                                }
                            }
                        }
                    }
                    
                    reporte.Text = "Estado: Guardando el documento Excel";

                    book.Save();
                    System.Threading.Thread.Sleep(1000);
                    app.Quit();

                    reporte.Text = "Estado: Programa finalizado";
                    MessageBox.Show("Programa finalizado");

                    Process.Start(saveRoute.Text);

                    System.Windows.Forms.Application.Exit();                    
                }
                else
                {
                    MessageBox.Show("El tipo de archivo seleccionado es incorrecto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("La ruta en donde guardar el archivo es incorrecta", "Ruta inexistente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            DialogResult dr = this.openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                foreach (String file in openFileDialog1.FileNames)
                {
                    nameFile.Text += file;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(Directory.Exists(saveRoute.Text))
            {                
                System.IO.File.Copy(System.IO.Path.GetFullPath(@"..\..\") + "archivos\\reemplazoCaracteres.bat", saveRoute.Text + "reemplazoCaracteres.bat", true);

                MessageBox.Show("Se ha guardado el archivo en la ruta definida.", "Alerta",
                       MessageBoxButtons.OK, MessageBoxIcon.Information);

                Process.Start(saveRoute.Text);
            }
            else
            {
                MessageBox.Show("El directorio para guardar los archivos no es válido.", "Alerta",
                       MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            button3_Click(sender, e);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string dummyFileName = "Save Here";

            SaveFileDialog sf = new SaveFileDialog();
            sf.CheckFileExists = false;
            sf.FileName = dummyFileName;
            sf.Filter = "Directory | directory";

            if (sf.ShowDialog() == DialogResult.OK)
            {
                string savePath = Path.GetDirectoryName(sf.FileName);
                saveRoute.Text = savePath + '\\';
            }
        }
    }
}
