using Net.SourceForge.Koogra;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace Per
{
    public partial class Form1 : Form
    {
        string my_file;

        readonly string[] stupeni = new string[] { "<10", "12-20", "22-30", "32-40", "42-50", "52-60", "62-70", "72-90", ">90" };
        //readonly List <string> stupeni = new List<string>() { "<10", "12-20", "22-30", "32-40", "42-50", "52-60", "62-70", "72-90", ">90" };

        public static int M_not_nul(int[] mas)
        {
            int col = 0;
            foreach (var i in mas)
            {
                if (i != 0)
                {
                    col++;
                }
            }
            return col;

        }

        public struct My_timber
        {
            public int[] all_arbor;
            public int[] lep_arbor;
            public int length_arbor;
            public int length_arbor_lep;

            public int[] all_trunk;
            public int[] lep_trunk;
            
        }

        //Количество диаметров, разделенных разделителями
        public static int[] Count_diametr(string input)
        {
            char[] my_separator = new char[] {'-', ','};
            //int[] count_diametr;
            //
            //считать все диаметры. 12-23-35
            string[] diametrs = input.Split(my_separator, StringSplitOptions.RemoveEmptyEntries);
            int length = diametrs.Length;
            int[] count_diametr = new int[length];
            for (int i = 0; i < length; ++i)
                count_diametr[i] = int.Parse(diametrs[i]);//последний диаметр
                        
            return count_diametr;
        }

        //точковка диаметров по ступеням толщины
        public static int[] Tochkovka(int[] arbor)
        {
            int[] tochkovka = new int[9];

            for (int i = 0; i < arbor.Length; ++i)
            {
                int diametr = arbor[i];
                if (diametr > 0 & diametr <= 10)
                {
                    ++tochkovka[0];
                }
                else if (diametr >10 & diametr <= 20)
                {
                    ++tochkovka[1];
                }
                else if (diametr > 20 & diametr <= 30)
                {
                    ++tochkovka[2];
                }
                else if (diametr > 30 & diametr <= 40)
                {
                    ++tochkovka[3];
                }
                else if (diametr > 40 & diametr <= 50)
                {
                    ++tochkovka[4];
                }
                else if (diametr > 50 & diametr <= 60)
                {
                    ++tochkovka[5];
                }
                else if (diametr > 60 & diametr <= 70)
                {
                    ++tochkovka[6];
                }
                else if (diametr > 70 & diametr <= 90)
                {
                    ++tochkovka[7];
                }
                else if (diametr > 90)
                {
                    ++tochkovka[8];
                }

            }
            return tochkovka;
        }


        //создает верифицированный список из номеров строк, которые проходят по критериям (дерево, на вырубку)
        public static List<uint> Verify_data(IWorksheet l_seets)
        {
            List<uint> data = new List<uint>();

            for (uint r = 7; r < l_seets.LastRow; ++r)
            {
                IRow row = l_seets.Rows.GetRow(r);
                if (row != null)
                {
                    if (row.GetCell(1) != null && row.GetCell(1).Value != null && row.GetCell(1).Value.ToString() == "Итого:")
                    {
                        break;
                    }
                    if (   row.GetCell(2) != null
                        && row.GetCell(2).Value != null
                        && row.GetCell(8) != null
                        && row.GetCell(8).Value != null
                        && row.GetCell(8).Value.ToString() == "Вырубить"
                        && row.GetCell(4) != null
                        && row.GetCell(4).Value != null
                        && row.GetCell(7) != null
                        && row.GetCell(7).Value != null)

                    {
                        data.Add(r);
                    }
                }

            }


            return data;
        }


        //создает массив строк (количество, диаметр, характеристика) при прохождении по номерам массива
        public static string[,] Read_xls_data(IWorksheet l_seets, List<uint> my_list)
        {
            //присваиваем переменной количество записей в списке валидных значений
            int count = my_list.Count;
            
            //Инициализируем двухмерный массив (количество строк как в списке, и фиксированно 3 столбца)
            string[,] data_xls = new string[count, 3];

            //Цикл проходит по списку и берет из xls данные из строки с номером из списка
            for (uint i = 0; i < count; ++i)
            {
                IRow row = l_seets.Rows.GetRow(my_list[(int)i]);

                data_xls[i, 0] = row.GetCell(2).Value.ToString();
                data_xls[i, 1] = row.GetCell(4).Value.ToString();
                data_xls[i, 2] = row.GetCell(7).Value.ToString();

            }

            return data_xls;

        }

        //переводит диапазон диаметров к одной записи в одномерном массиве
        public static My_timber Diametr(string[,] spisok)
        {
            int n = 0;
            int d = 0;
            My_timber timber = new My_timber();

            //вычисляет количество записей для создания массива
            for (int i = 0; i < spisok.GetUpperBound(0) + 1; ++i)
            {
                n += int.Parse(spisok[i, 0]);
            }
            //создает массив
            int[] diametr = new int[n];
            int[] diametr_lep = new int[n];

            //заполняет массив
            for (int i = 0; i < spisok.GetUpperBound(0) + 1; ++i)
            {
                //количество диаметров в записи (12-22) даст 2
                int[] count_d = Count_diametr(spisok[i, 1]);
                int length = count_d.Length;

                 //если одно дерево и один диаметр
                if (spisok[i, 0] == "1" && length == 1)
                {
                    diametr[d] = count_d[0];
                    if (spisok[i, 2].Contains("ЛЭП"))
                        {
                           diametr_lep[d] = count_d[0];
                        }
                    d++;
                }
                //если несколько деревьев и один диаметр
                else if (spisok[i, 0] != "1" && length == 1)
                {
                    int col = int.Parse(spisok[i, 0]);
                    for (int c = 0; c < col; ++c)
                    {
                        diametr[d] = count_d[0];
                        if (spisok[i, 2].Contains("ЛЭП"))
                        {
                            diametr_lep[d] = count_d[0];
                        }
                        d++;
                    }

                }
                //если одно дерево и несколько диаметров
                else if (spisok[i, 0] == "1" && length > 1)
                {
                    diametr[d] = count_d[length-1];
                    if (spisok[i, 2].Contains("ЛЭП"))
                    {
                        diametr_lep[d] = count_d[length-1];
                    }
                    d++;
                }
                //если несколько деревьев и несколько диаметров
                //
                //здесь добавить когда деревьев больше чем диаметров
                else if (spisok[i, 0] != "1" && length > 1)
                {
                    int der = length;
                    int col = int.Parse(spisok[i, 0]);
                    for (int c = col-1; c >= 0; --c)
                    {
                        if (der <1)
                        {
                            diametr[d] = count_d[0];
                            if (spisok[i, 2].Contains("ЛЭП"))
                            {
                                diametr_lep[d] = count_d[0];
                            }
                            der--;
                            d++;
                        }

                        else
                        {
                            diametr[d] = count_d[der - 1];
                            if (spisok[i, 2].Contains("ЛЭП"))
                            {
                                diametr_lep[d] = count_d[der - 1];
                            }
                            der--;
                            d++;

                        }

                    }

                }
            }

            timber.all_arbor = diametr;
            timber.lep_arbor = diametr_lep;
            timber.length_arbor = diametr.Length;
            timber.length_arbor_lep = M_not_nul(diametr_lep);
            return timber;
        }

        //Основная программа
        public void All_programm(string file_name)
        {
            // создать переменные для книги и первого листа
            IWorkbook book1 = WorkbookFactory.GetExcelBIFFReader(file_name);
            IWorksheet ws = book1.Worksheets.GetWorksheetByIndex(0);

            //если есть правильные данные
            if (Verify_data(ws).Count != 0)
            {
                //создаем массив данных
                var verify_rows = new List<uint>(Verify_data(ws));
                string[,] data_arbor = (Read_xls_data(ws, verify_rows));

                label6.Text = verify_rows.Count.ToString();
                
                //инициализация структуры и заполнение ее массивами диаметров
                My_timber timber = Diametr(data_arbor);
                //количество деревьев
                label5.Text = timber.length_arbor.ToString();
                label7.Text = timber.length_arbor.ToString();
                label8.Text = timber.length_arbor_lep.ToString();
                label9.Text = (timber.length_arbor - timber.length_arbor_lep).ToString();

                //удаляет все строки из таблицы
                dataGridView1.Rows.Clear();
                //вставляет в таблицу ступени толщины
                for (int i = 0; i < stupeni.Length; i++)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1[0, i].Value = stupeni[i];
                    dataGridView1[7, i].Value = Tochkovka(timber.all_arbor)[i];
                    dataGridView1[4, i].Value = Tochkovka(timber.lep_arbor)[i];
                    dataGridView1[1, i].Value = Tochkovka(timber.all_arbor)[i] - Tochkovka(timber.lep_arbor)[i];

                }

            }
        }

        public Form1()
        {
            InitializeComponent();
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dataGridView1.ColumnHeadersHeight = 30;
            dataGridView2.ColumnHeadersHeight = 30;

            //стиль для заголовка нижней таблицы
            dataGridView1.Columns[1].HeaderCell.Style.BackColor = Color.FromArgb(234, 240, 209);
            dataGridView1.Columns[1].Width = 85;
            dataGridView1.Columns[2].HeaderCell.Style.BackColor = Color.FromArgb(234, 240, 209);
            dataGridView1.Columns[2].Width = 85;
            dataGridView1.Columns[3].HeaderCell.Style.BackColor = Color.FromArgb(234, 240, 209);
            dataGridView1.Columns[3].Width = 85;

            dataGridView1.Columns[4].HeaderCell.Style.BackColor = Color.FromArgb(226, 240, 228);
            dataGridView1.Columns[4].Width = 85;
            dataGridView1.Columns[5].HeaderCell.Style.BackColor = Color.FromArgb(226, 240, 228);
            dataGridView1.Columns[5].Width = 85;
            dataGridView1.Columns[6].HeaderCell.Style.BackColor = Color.FromArgb(226, 240, 228);
            dataGridView1.Columns[6].Width = 85;

            dataGridView1.Columns[7].HeaderCell.Style.BackColor = Color.FromArgb(210, 235, 228);
            dataGridView1.Columns[7].Width = 85;
            dataGridView1.Columns[8].HeaderCell.Style.BackColor = Color.FromArgb(210, 235, 228);
            dataGridView1.Columns[8].Width = 85;
            dataGridView1.Columns[9].HeaderCell.Style.BackColor = Color.FromArgb(210, 235, 228);
            dataGridView1.Columns[9].Width = 85;




            //стиль для текста в нижней таблице
            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.FromArgb(255, 245, 229);
            //dataGridView1.Columns[0].DefaultCellStyle.Font = new Font("Arial", 10, FontStyle.Bold);  

            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.FromArgb(252, 255, 228);
            dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.FromArgb(252, 255, 228);
            dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.FromArgb(252, 255, 228);

            dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.FromArgb(236, 250, 238);
            dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.FromArgb(236, 250, 238);
            dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.FromArgb(236, 250, 238);

            dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.FromArgb(220, 245, 238);
            dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.FromArgb(220, 245, 238);
            dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.FromArgb(220, 245, 238);

            //стиль для заголовка верхней таблицы
            dataGridView2.Columns[1].HeaderCell.Style.BackColor = Color.FromArgb(234, 240, 215);
            dataGridView2.Columns[1].Width = 255;
            dataGridView2.Columns[2].HeaderCell.Style.BackColor = Color.FromArgb(226, 240, 228);
            dataGridView2.Columns[2].Width = 255;
            dataGridView2.Columns[3].HeaderCell.Style.BackColor = Color.FromArgb(210, 235, 228);
            dataGridView2.Columns[3].Width = 255;

            //стиль для заголовка ступеней толщины
            dataGridView1.Columns[0].HeaderCell.Style.BackColor = Color.FromArgb(244, 235, 219);
            dataGridView1.Columns[0].Width = 65;
            dataGridView2.Columns[0].HeaderCell.Style.BackColor = Color.FromArgb(244, 235, 219);
            dataGridView2.Columns[0].Width = 65;



        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Close();

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            // open dialog

            OpenFileDialog opendlg = new OpenFileDialog
            {
                Filter = "Xls Files (.xls)|*.xls"
            };

            if (DialogResult.OK == opendlg.ShowDialog())
            {
                // open xls file
                my_file = opendlg.FileName;
                label2.Text = my_file;
                All_programm(my_file);

            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void OpenFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] file = (string[])e.Data.GetData(DataFormats.FileDrop);
            my_file = file[0].ToString();
            label2.Text = my_file;
            // open xls file
            All_programm(my_file);


        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
               label2.Text = "Отпустите мышь";
               e.Effect = DragDropEffects.Copy;
            }
                
        }
    }
}
