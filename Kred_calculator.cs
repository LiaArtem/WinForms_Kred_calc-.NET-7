using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Linq;
using System.Xml.Linq;
using static System.Math;
using System.Xml;
using System.IO;

namespace Kred_calc
{
	public partial class Kred_calculator
	{
		public Kred_calculator()
		{
			InitializeComponent();
		}

		static int typ;
		static string[] file_path_ini_mas = Array.Empty<string>();
		static string[] type_ini_mas = Array.Empty<string>();
        public List<string> curr_code_date = new();
        public List<string> curr_code_val = new();
        static double p_sum_year;
        static Boolean p_is_year;
        static double p_sum_month;
        static Boolean p_is_month;
        static readonly string tec_kat = Application.StartupPath;

        // Пересчет дополнительных платежей
        #region Dop_plat

        private void Comiss_bank_TextChanged(System.Object sender, System.EventArgs e) {
            Dop_plat();
		}

		private void Comiss_strah1_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_strah2_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_strah3_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_notar1_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_notar2_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_notar3_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_notar4_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_notar5_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_rieltor1_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_rieltor2_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		private void Comiss_rieltor3_TextChanged(System.Object sender, System.EventArgs e) {
			Dop_plat();
		}

		// Расчет доп. платежей
		public void Dop_plat()
		{
			p_sum_year = 0;

            double p_calc = 0;
            p_calc += Dop_plat_in(this.comiss_bank.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.monthly_comiss_bank.Text, this.summa_ekv.Text, this.sum_kred.Text);
            this.bank_itog.Text = p_calc.ToString("#,0.00");
            //
            p_calc = 0;
            p_calc += Dop_plat_in(this.comiss_strah1.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.comiss_strah2.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.comiss_strah3.Text, this.summa_ekv.Text, this.sum_kred.Text);
            this.strax_itog.Text = p_calc.ToString("#,0.00");
            //
            p_calc = 0;
            p_calc += Dop_plat_in(this.comiss_notar1.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.comiss_notar2.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.comiss_notar3.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.comiss_notar4.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.comiss_notar5.Text, this.summa_ekv.Text, this.sum_kred.Text);
            this.notar_itog.Text = p_calc.ToString("#,0.00");

            //
            p_calc = 0;
            p_calc += Dop_plat_in(this.comiss_rieltor1.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.comiss_rieltor2.Text, this.summa_ekv.Text, this.sum_kred.Text);
            p_calc += Dop_plat_in(this.comiss_rieltor3.Text, this.summa_ekv.Text, this.sum_kred.Text);
            this.rieltor_itog.Text = p_calc.ToString("#,0.00");
            
			int sf = 0;
			if (double.Parse(bank_itog.Text) < 0) { sf = 1; bank_itog.Text = "#Error#"; }
            if (double.Parse(strax_itog.Text) < 0) { sf = 1; strax_itog.Text = "#Error#"; }
            if (double.Parse(notar_itog.Text) < 0) { sf = 1; notar_itog.Text = "#Error#"; }
			if (double.Parse(rieltor_itog.Text) < 0) { sf = 1; rieltor_itog.Text = "#Error#"; }			

			if (sf == 0) {
				this.sum_dop_plat.Text = 
                    ( double.Parse(bank_itog.Text)
                    + double.Parse(strax_itog.Text)
                    + double.Parse(notar_itog.Text) 
                    + double.Parse(rieltor_itog.Text)                     
                    ).ToString("#,0.00");
				//this.sum_na_ruki.Text = Prov_numeric(this.perv_vznos.Text).ToString("#,0.00");
			}
			else {
				this.sum_dop_plat.Text = "#Error#";
				//this.sum_na_ruki.Text = "#Error#";
			}
		}

        // Расчет месячных платежей
        public double Dop_plat_in_month(string t_sum_kred)
        {
            double sum_out = 0;
            string tt;
            tt = this.monthly_comiss_bank.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_strah1.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_strah2.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_strah3.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar1.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar2.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar3.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar4.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar5.Text;
            if (tt.Contains("%MONTH", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            return sum_out;
        }

        // Расчет годовых платежей
        public double Dop_plat_in_year(string t_sum_kred)
        {
            double sum_out = 0;
            string tt;
            tt = this.comiss_strah1.Text;
            if (tt.Contains("%YEAR", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_strah2.Text;
            if (tt.Contains("%YEAR", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_strah3.Text;
            if (tt.Contains("%YEAR", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar1.Text;
            if (tt.Contains("%YEAR", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar2.Text;
            if (tt.Contains("%YEAR", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar3.Text;
            if (tt.Contains("%YEAR", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar4.Text;
            if (tt.Contains("%YEAR", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            tt = this.comiss_notar5.Text;
            if (tt.Contains("%YEAR", StringComparison.CurrentCulture)) { sum_out += Dop_plat_in(tt, this.summa_ekv.Text, t_sum_kred); }
            return sum_out;
        }

        // Доп. платежи автоподстановки
        public static double Dop_plat_in(string t_in, string t_summa_ekv, string t_sum_kred)
		{
			string n = t_in;
			double s = Prov_numeric(t_summa_ekv);
			double s_kred = Prov_numeric(t_sum_kred);

            if (s < 0 || s_kred < 0)
            {
                return -1;
            }

            // оплата ежегодно
            p_is_year = false;
			if (n.Contains("%YEAR", StringComparison.CurrentCulture))
            {
				n = n.Replace("%YEAR", "");
				p_is_year = true;
			}

            // оплата ежемесяно
            p_is_month = false;
            if (n.Contains("%MONTH", StringComparison.CurrentCulture))
            {
                n = n.Replace("%MONTH", "");
                p_is_month = true;
            }

            double pl;
            double sum_out;
            // процент от суммы кредита
            if (n.Contains("%S", StringComparison.CurrentCulture))
            {
                n = n.Replace("%S", "");
                pl = Prov_numeric(n);
                if (s != 0)
                {
                    sum_out = Round((pl * s_kred) / 100, 2);
                }
                else
                {
                    sum_out = 0;
                }
            }
            // процент от суммы квартиры
            else if (n.Contains("%F", StringComparison.CurrentCulture))
            {
                n = n.Replace("%F", "");
                pl = Prov_numeric(n);
                if (s_kred != 0)
                {
                    sum_out = Round((pl * s) / 100, 2);
                }
                else
                {
                    sum_out = 0;
                }
            }
            else
            {
                // просто сумма
                sum_out = Prov_numeric(n);
            }

            if (p_is_year == true)
			{
				p_sum_year = sum_out + p_sum_year;
			}

            if (p_is_month == true)
            {
                p_sum_month = sum_out + p_sum_month;
            }

            return sum_out;
		}

#endregion

        // Прочие
        #region other

		// Поиск шаблонов
		public void Poisk_ini_files()
		{
			file_path_ini_mas = System.IO.Directory.GetFiles(tec_kat + "\\ini\\", "*.ini");
			type_ini_mas = new string[file_path_ini_mas.Length];
			for (var i = 0; i < file_path_ini_mas.Length; i++)
			{
				IniFile.IniFile_PATH(file_path_ini_mas[i]);
				type_ini_mas[i] = IniFile.IniReadValue("GLOBAL", "TYPE");
				this.type_rasch.Items.Add(IniFile.IniReadValue("GLOBAL", "NAME"));
			}
            IniFile.IniFile_PATH("");
		}

		// Чтение INI файла
		public void Read_ini_file_kred_calc()
		{
			try
			{
				IniFile.IniFile_PATH(file_path_ini_mas[this.type_rasch.SelectedIndex]);
				if (System.IO.File.Exists(file_path_ini_mas[this.type_rasch.SelectedIndex]) == true)
				{
                    this.priv_proc_stavka.Text = IniFile.IniReadValue("MAIN", "PRIV_PROC_STAVKA");
                    this.priv_srok_kred.Text = IniFile.IniReadValue("MAIN", "PRIV_SROK");
                    this.proc_stavka.Text = IniFile.IniReadValue("MAIN", "PROC_STAVKA");
                    this.summa.Text = IniFile.IniReadValue("MAIN", "SUMMA");
                    this.kurs.Text = IniFile.IniReadValue("MAIN", "KURS");
                    //this.summa_ekv.Text = IniFile.IniReadValue("MAIN", "SUMMA_EKV");
                    this.perv_vznos.Text = IniFile.IniReadValue("MAIN", "PERV_VZNOS");
                    this.proc_perv_vznos.Text = IniFile.IniReadValue("MAIN", "PERV_VZNOS_PROC");
                    this.srok_kred.Text = IniFile.IniReadValue("MAIN", "SROK");
                    this.date_cred.Value = DateTime.Today;
                    if (!string.IsNullOrEmpty(IniFile.IniReadValue("MAIN", "DATE_CRED")))
                    {
                        this.date_cred.Value = Convert.ToDateTime(IniFile.IniReadValue("MAIN", "DATE_CRED"));
                    }
                    if (IniFile.IniReadValue("MAIN", "TYPE_PROC") == "A")
                    {
                        this.type_proc.SelectedIndex = 1;
                    }
                    else if (IniFile.IniReadValue("MAIN", "TYPE_PROC") == "R")
                    {
                        this.type_proc.SelectedIndex = 2;
                    }
                    else
                    {
                        this.type_proc.SelectedIndex = 0;
                    }

                    this.kurs_start.Text = IniFile.IniReadValue("RASROCHKA", "KURS");
                    if (this.kurs_start.Text == "") { this.kurs_start.Text = "1"; } 
                    this.kurs_year_0.Text = IniFile.IniReadValue("RASROCHKA", "KURS_YEAR_0");
                    if (this.kurs_year_0.Text == "") { this.kurs_year_0.Text = "1"; }
                    this.kurs_year_1.Text = IniFile.IniReadValue("RASROCHKA", "KURS_YEAR_1");
                    if (this.kurs_year_1.Text == "") { this.kurs_year_1.Text = "1"; }
                    this.kurs_year_2.Text = IniFile.IniReadValue("RASROCHKA", "KURS_YEAR_2");
                    if (this.kurs_year_2.Text == "") { this.kurs_year_2.Text = "1"; }
                    this.kurs_year_3.Text = IniFile.IniReadValue("RASROCHKA", "KURS_YEAR_3");
                    if (this.kurs_year_3.Text == "") { this.kurs_year_3.Text = "1"; }
                    this.kurs_year_4.Text = IniFile.IniReadValue("RASROCHKA", "KURS_YEAR_4");
                    if (this.kurs_year_4.Text == "") { this.kurs_year_4.Text = "1"; }

                    this.year_0.Text = this.date_cred.Value.Year.ToString();
                    this.year_1.Text = this.date_cred.Value.AddYears(1).Year.ToString();
                    this.year_2.Text = this.date_cred.Value.AddYears(2).Year.ToString();
                    this.year_3.Text = this.date_cred.Value.AddYears(3).Year.ToString();
                    this.year_4.Text = this.date_cred.Value.AddYears(4).Year.ToString();

                    this.curr_code.Text = "UAH";
                    if (!string.IsNullOrEmpty(IniFile.IniReadValue("MAIN", "CURR_CODE")))
                    {
                        this.curr_code.Text = IniFile.IniReadValue("MAIN", "CURR_CODE");
                    }
                    if (this.curr_code.Text == "UAH") {
                        this.kurs.Text = "1";
                        this.kurs.Enabled = false;
                    }
                    else {
                        this.kurs.Enabled = true;
                        if (!string.IsNullOrEmpty(IniFile.IniReadValue("MAIN", "KURS")))
                        {
                            this.kurs.Text = IniFile.IniReadValue("MAIN", "KURS");
                        }
                        else
                        {
                            this.kurs.Text = Get_kurs_nbu_in_site(this.curr_code.Text, this.date_cred.Value);
                        }

                    }

                    // квартира
                    if (type_ini_mas[this.type_rasch.SelectedIndex] == "K")
					{
                        this.comiss_bank.Text = IniFile.IniReadValue("DOPOLN", "BANK_KOMISS_OBSLUGV");
                        this.monthly_comiss_bank.Text = IniFile.IniReadValue("DOPOLN", "BANK_KOMISS_MONTHLY");
                        this.comiss_strah1.Text = IniFile.IniReadValue("DOPOLN", "STRA_KOMISS_IPOTEKA");
						this.comiss_strah2.Text = IniFile.IniReadValue("DOPOLN", "STRA_KOMISS_SOBSTVN");
						this.comiss_strah3.Text = IniFile.IniReadValue("DOPOLN", "STRA_KOMISS_NESLUCH");
						this.comiss_notar1.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_IREESTR");
						this.comiss_notar2.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_DOG_ZAL");
						this.comiss_notar3.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_POSHINA");
						this.comiss_notar4.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_PENFOND");
						this.comiss_notar5.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_DOG_POK");
						this.comiss_rieltor1.Text = IniFile.IniReadValue("DOPOLN", "RIEL_KOMISS_EKSPERT");
						this.comiss_rieltor2.Text = IniFile.IniReadValue("DOPOLN", "RIEL_KOMISS_KONSULT");
						this.comiss_rieltor3.Text = IniFile.IniReadValue("DOPOLN", "RIEL_KOMISS_REG_BTI");
					}
                    // автомобиль
					else
					{
						this.comiss_bank.Text = IniFile.IniReadValue("DOPOLN", "BANK_KOMISS_OBSLUGV");
                        this.monthly_comiss_bank.Text = IniFile.IniReadValue("DOPOLN", "BANK_KOMISS_MONTHLY");
                        this.comiss_strah1.Text = IniFile.IniReadValue("DOPOLN", "STRA_KOMISS_KASKO");
						this.comiss_strah2.Text = IniFile.IniReadValue("DOPOLN", "STRA_KOMISS_OTHER");
						this.comiss_strah3.Text = IniFile.IniReadValue("DOPOLN", "STRA_KOMISS_OSAGO");
						this.comiss_notar1.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_IREESTR");
						this.comiss_notar2.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_DOG_ZAL");
						this.comiss_notar3.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_POSHINA");
						this.comiss_notar4.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_PENFOND");
						this.comiss_notar5.Text = IniFile.IniReadValue("DOPOLN", "NOTA_KOMISS_DOG_POK");
						this.comiss_rieltor1.Text = IniFile.IniReadValue("DOPOLN", "RIEL_KOMISS_TRANSPORT");
						this.comiss_rieltor2.Text = IniFile.IniReadValue("DOPOLN", "RIEL_KOMISS_NOTAR");
						this.comiss_rieltor3.Text = IniFile.IniReadValue("DOPOLN", "RIEL_KOMISS_GAI");
					}

					if (string.IsNullOrEmpty(this.proc_perv_vznos.Text))
					{
						this.CheckBox1.Checked = false;
					}
					else
					{
						this.CheckBox1.Checked = true;
					}
				}
				else
				{
					MessageBox.Show("Отсутствует файл по адресу:" + tec_kat + "\\kred_calc.ini", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
			}
			catch
			{
				MessageBox.Show("Ошибка в файле по адресу:" + tec_kat + "\\kred_calc.ini", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
		}

        // Загрузка описаний доп. параметров
		public void Load_radio()
		{
			if (type_ini_mas[this.type_rasch.SelectedIndex] == "K")
			{
				// Квартира
				//this.Label2.Text = "Стоимость квартиры";
				//
				this.Label22.Text = "%F - % от стоимости квартиры";
				//
				this.Label9.Text = "Страхование предмета ипотеки";
				this.Label10.Text = "По договору страхования жизни";
				this.Label11.Text = "Страхование от нещасного случая";
				//
				this.TabPage4.Text = "Риелтор";
				this.Label20.Text = "Экспертная оценка недвижимости";
				this.Label18.Text = "Консультация и оформл. документов";
				this.Label19.Text = "Регистрация договора в БТИ";
			}
			else
			{
				// Авто
				//this.Label2.Text = "Стоимость авто";
				//
				this.Label22.Text = "%F - % от стоимости авто";
				//
				this.Label9.Text = "КАСКО";
				this.Label10.Text = "Страхование жизни";
				this.Label11.Text = "ОСАГО";
				//
				this.TabPage4.Text = "Прочие";
				this.Label20.Text = "Транспортный сбор";
				this.Label18.Text = "Услуги нотариуса";
				this.Label19.Text = "Регистрация в ГАИ";
			}
		}

		// Проверка текстовых сообщений на число
		private static double Prov_numeric(string text_in)
		{
            double n;
            if (double.TryParse(text_in, out _) == false)
			{
				n = 0;
				if (!string.IsNullOrEmpty(text_in))
				{
					n = -1;
				}
			}
			else
			{
				n = Round(double.Parse(text_in), 2);
			}
			return n;
		}

        // Расчет и вывод таблицы
        public void Paint_table()
        {
            this.DataGridView1.Columns.Clear();
            this.DataGridView1.RowHeadersWidth = 10;
            this.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DataGridView1.ReadOnly = true;
            this.DataGridView1.AllowUserToAddRows = false;
            this.DataGridView1.AllowUserToDeleteRows = false;
            // Добавление колонок
            //
            DataGridViewTextBoxColumn com1 = new()
            {
                HeaderText = "Дата (мес./год)",
                ReadOnly = true,
                Width = 70
            };
            this.DataGridView1.Columns.Add(com1);
            //
            DataGridViewTextBoxColumn com2 = new()
            {
                HeaderText = "Задолжность по кредиту",
                ReadOnly = true,
                Width = 80
            };
            this.DataGridView1.Columns.Add(com2);
            //
            DataGridViewTextBoxColumn com3 = new()
            {
                HeaderText = "Платеж по процентам",
                ReadOnly = true,
                Width = 80
            };
            this.DataGridView1.Columns.Add(com3);
            //
            DataGridViewTextBoxColumn com4 = new()
            {
                HeaderText = "Платежи, кредит",
                ReadOnly = true,
                Width = 80
            };
            this.DataGridView1.Columns.Add(com4);
            //
            DataGridViewTextBoxColumn com5 = new()
            {
                HeaderText = "Переплата",
                ReadOnly = true,
                Width = 80
            };
            this.DataGridView1.Columns.Add(com5);
            //
            DataGridViewTextBoxColumn com6 = new()
            {
                HeaderText = "Платежи, дополн.",
                ReadOnly = true,
                Width = 80
            };
            this.DataGridView1.Columns.Add(com6);
            //
            DataGridViewTextBoxColumn com7 = new()
            {
                HeaderText = "Общий платеж",
                ReadOnly = true,
                Width = 80
            };
			this.DataGridView1.Columns.Add(com7);

            this.DataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.DataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.DataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.DataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.DataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.DataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        // Расчет графика
        public void Rashet_formula(string type_rashet, string type_annuitet)
		{

			int i = 0;
			this.DataGridView1.Rows.Clear();
			double sum_kred = Prov_numeric(this.sum_kred.Text);
			double proc_stavka = Prov_numeric(this.proc_stavka.Text);            
            double priv_proc_stavka = Prov_numeric(this.priv_proc_stavka.Text);
            int srok = Convert.ToInt32(Prov_numeric(this.srok_kred.Text));
            int priv_srok = Convert.ToInt32(Prov_numeric(this.priv_srok_kred.Text));            

            if (Prov_numeric(this.sum_kred.Text) < 0)
			{
				MessageBox.Show("Расчет и вывод графика невозможен !!! Есть нечисловые значения СУММА КРЕДИТА!!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			if (Prov_numeric(this.sum_dop_plat.Text) < 0)
			{
				MessageBox.Show("Расчет и вывод графика невозможен !!! Есть нечисловые значения СУММА ДОПОЛНИТЕЛЬНЫХ ПЛАТЕЖЕЙ!!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			if (Prov_numeric(this.proc_stavka.Text) < 0)
			{
				MessageBox.Show("Расчет и вывод графика невозможен !!! Есть нечисловые значения ПРОЦЕНТНАЯ СТАВКА!!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			if (Prov_numeric(this.srok_kred.Text) < 0)
			{
				MessageBox.Show("Расчет и вывод графика невозможен !!! Есть нечисловые значения СРОК КРЕДИТА!!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			double sum_one = 0;
			if (on_dop_rasx.Checked == true)
			{
				if (dop_rasx_month_or_all.Checked == true)
				{
					sum_one = Prov_numeric(this.sum_dop_plat.Text);
				}
				else
				{
					sum_kred += Prov_numeric(this.sum_dop_plat.Text);
				}
			}

            // Аннуитет
			if (type_rashet == "аннуитетная")
			{
				// расчет кредитного портфеля
				int zn = 0;
				int zc = 0;
				// Расчитываем процентную ставку выраженную в долях
				if (type_annuitet == "30/360")
				{
					zc = 30;
					zn = 360;
				}
				else if (type_annuitet == "факт/360")
				{
					zc = Public.LastDayOfMonth(DateTime.Today);
					zn = 360;
				}
				else if (type_annuitet == "факт/факт")
				{
					zc = Public.LastDayOfMonth(DateTime.Today);
					zn = Public.KolDayOfYear(DateTime.Today);
				}

                //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                // Без льготного периода
                if (priv_srok == 0)
                {
                    proc_stavka = (proc_stavka * 0.01 / zn) * zc;
                    // Сумма аннуитетного платежа
                    double annuitet = sum_kred * proc_stavka / (1 - Math.Pow(1 + proc_stavka, -1 * srok));
                    // Переплата по кредиту
                    double sum_plat = Prov_numeric(this.sum_plat.Text);
                    double sum_pereplata = 0;
                    if (sum_plat > annuitet)
                    {
                        sum_pereplata = sum_plat - annuitet;
                        annuitet = sum_plat;
                    }

                    //
                    double summ = sum_kred;
                    DateTime d_date = this.date_cred.Value;
                    double summ_pro = sum_kred * proc_stavka;
                    double n_pr = 0;
                    double n_ob = annuitet * srok + sum_one;
                    double summ_dop = 0;
                    int srok_new = 0;

                    for (i = 1; i <= srok; i++)
                    {
                        // учет ежегодных
                        double sum_year = 0;
                        sum_year = 0;
                        if ((i - 1) % 12 == 0 && i != 1)
                        {
                            sum_year = Dop_plat_in_year(summ.ToString());
                            //sum_year = p_sum_year;
                        }
                        // учет ежемесяных
                        double sum_month = Dop_plat_in_month(summ.ToString());

                        //
                        this.DataGridView1.Rows.Add(new object[]
                        {
                        Get_date_month(d_date),
                        Round(summ, 2).ToString("#,0.00"),
                        Round(summ_pro, 2).ToString("#,0.00"),
                        (Round(annuitet - summ_pro, 2)).ToString("#,0.00"),
                        (Round(sum_pereplata, 2)).ToString("#,0.00"),
                        (Round(sum_one + sum_year + sum_month, 2)).ToString("#,0.00"),
                        (Round(annuitet + sum_one + sum_year + sum_month, 2)).ToString("#,0.00")
                        });
                        d_date = d_date.AddMonths(1);
                        summ = summ - annuitet + proc_stavka * summ;
                        n_pr = n_pr + summ_pro + sum_one;
                        summ_pro = summ * proc_stavka;
                        summ_dop = summ_dop + sum_one + sum_year + sum_month;
                        sum_one = 0;
                        if (summ < 0) { break; }
                        srok_new++;
                    }

                    this.srok_cred_new.Text = srok_new.ToString("#");
                    this.srok_kred_year_new.Text = (Round(Prov_numeric(this.srok_cred_new.Text) / 12, 2)).ToString("#,0.00");

                    // Итого
                    this.DataGridView1.Rows.Add(new object[] { "Итого:", "", Round(n_pr, 2).ToString("#,0.00"), Round(sum_kred, 2).ToString("#,0.00"), "", Round(summ_dop, 2).ToString("#,0.00"), Round(n_ob, 2).ToString("#,0.00") });
                    // Переплата
                    this.DataGridView1.Rows.Add(new object[] { "Переплата:", "", "", "", "", "", Round(n_pr, 2).ToString("#,0.00") });
                    this.pereplata.Text = Round(n_pr + summ_dop, 2).ToString("#,0.00");
                }
                //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                // С льготным периодом
                else
                {
                    if (priv_proc_stavka < 0.01) { priv_proc_stavka = 0.01; }
                    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    // Льготный период
                        priv_proc_stavka = (priv_proc_stavka * 0.01 / zn) * zc;
                    // Сумма аннуитетного платежа
                    double annuitet = sum_kred * priv_proc_stavka / (1 - Math.Pow(1 + priv_proc_stavka, -1 * srok));
                    
                    //
                    double summ = sum_kred;
                    DateTime d_date = this.date_cred.Value;
                    double summ_pro = sum_kred * priv_proc_stavka;
                    double n_pr = 0;
                    double n_ob = annuitet * srok + sum_one;
                    double summ_dop = 0;
                    for (i = 1; i <= priv_srok; i++)
                    {
                        // учет ежегодных
                        double sum_year = 0;
                        sum_year = 0;
                        if ((i - 1) % 12 == 0 && i != 1)
                        {
                            sum_year = Dop_plat_in_year(summ.ToString());
                            //sum_year = p_sum_year;
                        }
                        // учет ежемесяных
                        double sum_month = Dop_plat_in_month(summ.ToString());

                        //
                        this.DataGridView1.Rows.Add(new object[]
                        {
                        Get_date_month(d_date),
                        Round(summ, 2).ToString("#,0.00"),
                        Round(summ_pro, 2).ToString("#,0.00"),
                        (Round(annuitet - summ_pro, 2)).ToString("#,0.00"),
                        "0.00",
                        (Round(sum_one + sum_year + sum_month, 2)).ToString("#,0.00"),
                        (Round(annuitet + sum_one + sum_year + sum_month, 2)).ToString("#,0.00")
                        });
                        d_date = d_date.AddMonths(1);
                        summ = summ - annuitet + priv_proc_stavka * summ;
                        n_pr = n_pr + summ_pro + sum_one;
                        summ_pro = summ * priv_proc_stavka;
                        summ_dop = summ_dop + sum_one + sum_year + sum_month;
                        sum_one = 0;
                        if (summ < 0) { break; }
                    }

                    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    // Обычный период
                    srok -= priv_srok;
                    proc_stavka = (proc_stavka * 0.01 / zn) * zc;
                    // Сумма аннуитетного платежа
                    annuitet = summ * proc_stavka / (1 - Math.Pow(1 + proc_stavka, -1 * srok));
                    //                                        
                    summ_pro = summ * proc_stavka;
                    n_ob += (annuitet * srok);
                    for (i = 1; i <= srok; i++)
                    {
                        // учет ежегодных
                        double sum_year = 0;
                        sum_year = 0;
                        if ((i - 1) % 12 == 0 && i != 1)
                        {
                            sum_year = Dop_plat_in_year(summ.ToString());
                            //sum_year = p_sum_year;
                        }
                        // учет ежемесяных
                        double sum_month = Dop_plat_in_month(summ.ToString());

                        //
                        this.DataGridView1.Rows.Add(new object[]
                        {
                        Get_date_month(d_date),
                        Round(summ, 2).ToString("#,0.00"),
                        Round(summ_pro, 2).ToString("#,0.00"),
                        (Round(annuitet - summ_pro, 2)).ToString("#,0.00"),
                        "0.00",
                        (Round(sum_one + sum_year + sum_month, 2)).ToString("#,0.00"),
                        (Round(annuitet + sum_one + sum_year + sum_month, 2)).ToString("#,0.00")
                        });
                        d_date = d_date.AddMonths(1);
                        summ = summ - annuitet + proc_stavka * summ;
                        n_pr = n_pr + summ_pro + sum_one;
                        summ_pro = summ * proc_stavka;
                        summ_dop = summ_dop + sum_one + sum_year + sum_month;
                        sum_one = 0;
                        if (summ < 0) { break; }
                    }

                    // Итого
                    this.DataGridView1.Rows.Add(new object[] { "Итого:", "", Round(n_pr, 2).ToString("#,0.00"), Round(sum_kred, 2).ToString("#,0.00"), "", Round(summ_dop, 2).ToString("#,0.00"), Round(n_ob, 2).ToString("#,0.00") });
                    // Переплата
                    this.DataGridView1.Rows.Add(new object[] { "Переплата:", "", "", "", "", "", Round(n_pr, 2).ToString("#,0.00") });
                    this.pereplata.Text = Round(n_pr + summ_dop, 2).ToString("#,0.00");
                }
            }
			//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            // Стандартный
			if (type_rashet == "классика")
			{				
				double summ = sum_kred;
                double summ_graf = sum_kred;
                double n_pr = 0;
				double n_ob = 0;
                double n_cred = 0;
                double n_perepl = 0;
                double sum_year = 0;
                double pr;
                double summ_dop = 0;
                double sum_plat = Prov_numeric(this.sum_plat.Text);
                string[] mass_date = new string[srok];
                double[,] mass_num = new double[6, srok];
                double zc = 0;
                double zn = 0;
                double sum_pereplata = 0;
                int srok_new = 0;

                // платежи кредит
                DateTime d_date = this.date_cred.Value;
				for (i = 1; i <= srok; i++)
				{                    
                    // Расчитываем процентную ставку выраженную в долях
                    if (type_annuitet == "30/360") {
                        zc = 30; zn = 360;
                    }
                    else if (type_annuitet == "факт/360") {
                        zc = Public.LastDayOfMonth(d_date); zn = 360;                        
                    }
                    else if (type_annuitet == "факт/факт") {
                        zc = Public.LastDayOfMonth(d_date); zn = Public.KolDayOfYear(d_date);                        
                    }

                    // льготная
                    if (i <= priv_srok) {
                        pr = summ_graf * priv_proc_stavka * (zc/zn) / 100;
                    }
                    // обычная
                    else {
                        pr = summ_graf * proc_stavka * (zc / zn) / 100;
                    }
                     
					// учет ежегодных
					sum_year = 0;
					if ((i - 1) % 12 == 0 && i != 1)
					{
                        sum_year = sum_year = Dop_plat_in_year(summ_graf.ToString());
                        //sum_year = p_sum_year;
                    }
                    // учет ежемесяных
                    double sum_month = Dop_plat_in_month(summ_graf.ToString());
                    // учет переплаты
                    double calc_sum_cred = sum_kred / srok;
                    double sum_itog = Round(calc_sum_cred + pr + sum_one + sum_year + sum_month, 2);

                    mass_date[i - 1] = Get_date_month(d_date);
                    mass_num[0, i - 1] = Round(summ, 2);
                    mass_num[1, i - 1] = Round(pr, 2);
                    mass_num[2, i - 1] = Round(calc_sum_cred, 2);
                    mass_num[3, i - 1] = Round(sum_one + sum_year + sum_month, 2);                    
                    mass_num[5, i - 1] = 0;

                    if (sum_plat > Round(calc_sum_cred + pr, 2))
                    {
                        sum_pereplata = sum_plat - Round(calc_sum_cred + pr, 2); // переплата
                        mass_num[5, i - 1] = sum_pereplata;
                        // если последний платеж, корректируем переплату
                        if (summ - (sum_plat - Round(pr, 2)) <= 0)
                        {
                            sum_pereplata = 0;
                            calc_sum_cred = summ;
                            mass_num[2, i - 1] = calc_sum_cred;
                            mass_num[5, i - 1] = sum_pereplata;
                            // пересчет %
                            // Расчитываем процентную ставку выраженную в долях
                            if (type_annuitet == "30/360") {
                                zc = 30; zn = 360;
                            }
                            else if (type_annuitet == "факт/360") {
                                zc = Public.LastDayOfMonth(d_date); zn = 360;
                            }
                            else if (type_annuitet == "факт/факт") {
                                zc = Public.LastDayOfMonth(d_date); zn = Public.KolDayOfYear(d_date);
                            }

                            // льготная
                            if (i <= priv_srok) {
                                pr = summ * priv_proc_stavka * (zc / zn) / 100;
                            }
                            // обычная
                            else {
                                pr = summ * proc_stavka * (zc / zn) / 100;
                            }

                            mass_num[1, i - 1] = Round(pr, 2);
                            sum_itog = Round(calc_sum_cred + pr + sum_one + sum_year + sum_month, 2);
                            ///////////////////////////////////////////////////////////////////////////
                        }
                        summ -= (sum_plat - Round(pr, 2));
                    }
                    else
                    {
                        summ -= Round(calc_sum_cred, 2);
                    }

                    mass_num[4, i - 1] = sum_itog + sum_pereplata;

                    summ_graf -= Round(calc_sum_cred, 2);
                    d_date = d_date.AddMonths(1);					 
					n_pr += pr;
					n_ob = n_ob + calc_sum_cred + pr + sum_one + sum_pereplata;
                    n_cred += calc_sum_cred;
                    n_perepl += sum_pereplata;
                    summ_dop = summ_dop + sum_one + sum_year + sum_month;
                    sum_one = 0;
                    if (summ < 0) { break; } 
                }

                //
                srok_new = 0;
                for (i = 1; i <= srok; i++)
                {
                    if (mass_num[2, i - 1] == 0) { break; }

                    this.DataGridView1.Rows.Add(new object[]
                    {                        
                        mass_date[i - 1],
                        mass_num[0, i - 1].ToString("#,0.00"),
                        mass_num[1, i - 1].ToString("#,0.00"),
                        mass_num[2, i - 1].ToString("#,0.00"),
                        mass_num[5, i - 1].ToString("#,0.00"),
                        mass_num[3, i - 1].ToString("#,0.00"),
                        mass_num[4, i - 1].ToString("#,0.00")
                    });
                    srok_new++;
                }

                this.srok_cred_new.Text = srok_new.ToString("#");
                this.srok_kred_year_new.Text = (Round(Prov_numeric(this.srok_cred_new.Text) / 12, 2)).ToString("#,0.00");

                // Итого
                this.DataGridView1.Rows.Add(new object[] {"Итого:", "", Round(n_pr, 2).ToString("#,0.00"),
                                                                        Round(n_cred, 2).ToString("#,0.00"),
                                                                        Round(n_perepl, 2).ToString("#,0.00"),
                                                                        Round(summ_dop, 2).ToString("#,0.00"),
                                                                        Round(n_ob, 2).ToString("#,0.00") });
                // Переплата
                this.DataGridView1.Rows.Add(new object[] { "Переплата:", "", "", "", "", "", Round(n_pr + summ_dop, 2).ToString("#,0.00") });
                this.pereplata.Text = Round(n_pr + summ_dop, 2).ToString("#,0.00");
            }
            //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            // Рассрочка
            if (type_rashet == "рассрочка")
            {
                double summ = sum_kred;
                double summ_graf = sum_kred;
                double n_pr = 0;
                double n_ob = 0;
                double n_cred = 0;
                double n_perepl = 0;
                double sum_year = 0;
                double pr;
                double summ_dop = 0;
                double sum_plat = Prov_numeric(this.sum_plat.Text);
                string[] mass_date = new string[srok];
                double[,] mass_num = new double[6, srok];
                double sum_pereplata = 0;
                int srok_new = 0;
                double kurs_start = Prov_numeric(this.kurs_start.Text);
                double kurs_year_0 = Prov_numeric(this.kurs_year_0.Text);
                double kurs_year_1 = Prov_numeric(this.kurs_year_1.Text);
                double kurs_year_2 = Prov_numeric(this.kurs_year_2.Text);
                double kurs_year_3 = Prov_numeric(this.kurs_year_3.Text);
                double kurs_year_4 = Prov_numeric(this.kurs_year_4.Text);

                // платежи кредит
                DateTime d_date = this.date_cred.Value;
                for (i = 1; i <= srok; i++)
                {
                    // начальный год
                    if (d_date.Year == this.date_cred.Value.Year) {
                        pr = ((kurs_year_0 / kurs_start) - 1) * (sum_kred / srok);
                    }
                    else if (d_date.Year == this.date_cred.Value.AddYears(1).Year) {
                        pr = ((kurs_year_1 / kurs_start) - 1) * (sum_kred / srok);
                    }
                    else if (d_date.Year == this.date_cred.Value.AddYears(2).Year)
                    {
                        pr = ((kurs_year_2 / kurs_start) - 1) * (sum_kred / srok);
                    }
                    else if (d_date.Year == this.date_cred.Value.AddYears(3).Year)
                    {
                        pr = ((kurs_year_3 / kurs_start) - 1) * (sum_kred / srok);
                    }
                    else if (d_date.Year == this.date_cred.Value.AddYears(4).Year || d_date.Year > this.date_cred.Value.AddYears(4).Year)
                    {
                        pr = ((kurs_year_4 / kurs_start) - 1) * (sum_kred / srok);
                    }
                    else { pr = 0; }

                    // учет ежегодных
                    sum_year = 0;
                    if ((i - 1) % 12 == 0 && i != 1)
                    {
                        sum_year = sum_year = Dop_plat_in_year(summ_graf.ToString());
                        //sum_year = p_sum_year;
                    }
                    // учет ежемесяных
                    double sum_month = Dop_plat_in_month(summ_graf.ToString());
                    // учет переплаты
                    double calc_sum_cred = sum_kred / srok;
                    double sum_itog = Round(calc_sum_cred + pr + sum_one + sum_year + sum_month, 2);

                    mass_date[i - 1] = Get_date_month(d_date);
                    mass_num[0, i - 1] = Round(summ, 2);
                    mass_num[1, i - 1] = Round(pr, 2);
                    mass_num[2, i - 1] = Round(calc_sum_cred, 2);
                    mass_num[3, i - 1] = Round(sum_one + sum_year + sum_month, 2);
                    mass_num[5, i - 1] = 0;

                    if (sum_plat > Round(calc_sum_cred + pr, 2))
                    {
                        sum_pereplata = sum_plat - Round(calc_sum_cred + pr, 2); // переплата
                        mass_num[5, i - 1] = sum_pereplata;
                        // если последний платеж, корректируем переплату
                        if (summ - (sum_plat - Round(pr, 2)) <= 0)
                        {
                            sum_pereplata = 0;
                            calc_sum_cred = summ;
                            mass_num[2, i - 1] = calc_sum_cred;
                            mass_num[5, i - 1] = sum_pereplata;
                            // пересчет %

                            // начальный год
                            if (d_date.Year == this.date_cred.Value.Year)
                            {
                                pr = (kurs_year_0 / kurs_start) * calc_sum_cred;
                            }
                            else if (d_date.Year == this.date_cred.Value.AddYears(1).Year)
                            {
                                pr = (kurs_year_1 / kurs_start) * calc_sum_cred;
                            }
                            else if (d_date.Year == this.date_cred.Value.AddYears(2).Year)
                            {
                                pr = (kurs_year_2 / kurs_start) * calc_sum_cred;
                            }
                            else if (d_date.Year == this.date_cred.Value.AddYears(3).Year)
                            {
                                pr = (kurs_year_3 / kurs_start) * calc_sum_cred;
                            }
                            else if (d_date.Year == this.date_cred.Value.AddYears(4).Year || d_date.Year > this.date_cred.Value.AddYears(4).Year)
                            {
                                pr = (kurs_year_4 / kurs_start) * calc_sum_cred;
                            }
                            else { pr = 0; }

                            mass_num[1, i - 1] = Round(pr, 2);
                            sum_itog = Round(calc_sum_cred + pr + sum_one + sum_year + sum_month, 2);
                            ///////////////////////////////////////////////////////////////////////////
                        }
                        summ -= (sum_plat - Round(pr, 2));
                    }
                    else
                    {
                        summ -= Round(calc_sum_cred, 2);
                    }

                    mass_num[4, i - 1] = sum_itog + sum_pereplata;

                    summ_graf -= Round(calc_sum_cred, 2);
                    d_date = d_date.AddMonths(1);
                    n_pr += pr;
                    n_ob = n_ob + calc_sum_cred + pr + sum_one + sum_pereplata;
                    n_cred += calc_sum_cred;
                    n_perepl += sum_pereplata;
                    summ_dop = summ_dop + sum_one + sum_year + sum_month;
                    sum_one = 0;
                    if (summ < 0) { break; }
                }

                //
                srok_new = 0;
                for (i = 1; i <= srok; i++)
                {
                    if (mass_num[2, i - 1] == 0) { break; }

                    this.DataGridView1.Rows.Add(new object[]
                    {
                        mass_date[i - 1],
                        mass_num[0, i - 1].ToString("#,0.00"),
                        mass_num[1, i - 1].ToString("#,0.00"),
                        mass_num[2, i - 1].ToString("#,0.00"),
                        mass_num[5, i - 1].ToString("#,0.00"),
                        mass_num[3, i - 1].ToString("#,0.00"),
                        mass_num[4, i - 1].ToString("#,0.00")
                    });
                    srok_new++;
                }

                this.srok_cred_new.Text = srok_new.ToString("#");
                this.srok_kred_year_new.Text = (Round(Prov_numeric(this.srok_cred_new.Text) / 12, 2)).ToString("#,0.00");

                // Итого
                this.DataGridView1.Rows.Add(new object[] {"Итого:", "", Round(n_pr, 2).ToString("#,0.00"),
                                                                        Round(n_cred, 2).ToString("#,0.00"),
                                                                        Round(n_perepl, 2).ToString("#,0.00"),
                                                                        Round(summ_dop, 2).ToString("#,0.00"),
                                                                        Round(n_ob, 2).ToString("#,0.00") });
                // Переплата
                this.DataGridView1.Rows.Add(new object[] { "Переплата:", "", "", "", "", "", Round(n_pr + summ_dop, 2).ToString("#,0.00") });
                this.pereplata.Text = Round(n_pr + summ_dop, 2).ToString("#,0.00");
            }

            // Проставляем цвет
            if (this.DataGridView1.Rows.Count >= 3)
            {
                i = 0;
                typ = 0;
                string? etalon = DataGridView1.Rows[i].Cells[0].Value.ToString();
                if (etalon != null) { etalon = etalon.ToString()[..4]; }

                Get_color(i);
                for (i = 1; i <= this.DataGridView1.Rows.Count - 1 - 2; i++)
                {
                    string? prov = DataGridView1.Rows[i].Cells[0].Value.ToString();
                    if (prov != null) { prov = prov.ToString()[..4]; }
                    if (prov == etalon)
                    {
                        Get_color(i);
                    }
                    else
                    {
                        typ++;
                        if (typ > 1)
                        {
                            typ = 0;
                        }                        
                        etalon = DataGridView1.Rows[i].Cells[0].Value.ToString();
                        if (etalon != null) { etalon = etalon.ToString()[..4]; }

                        Get_color(i);
                    }
                }
            
			// Прошедшие месяцы
			//for (i = 0; i <= this.DataGridView1.Rows.Count - 1 - 2; i++)
			//{
			//	typ = 2;
				//if (new DateTime(Convert.ToInt32(this.DataGridView1.Rows[i].Cells[0].Value.ToString().Substring(0, 4)), Convert.ToInt32(this.DataGridView1.Rows[i].Cells[0].Value.ToString().Substring(5, 2)), this.date_cred.Value.Day) < DateTime.Today)
				//{
				//	Get_color(i);
				//}
			//}
			this.DataGridView1.Rows[DataGridView1.Rows.Count - 1 - 1].DefaultCellStyle.BackColor = Color.LightGreen;
			this.DataGridView1.Rows[DataGridView1.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightBlue;
            }
        }

		public void Get_color(int i)
		{
			if (typ == 0)
			{
				this.DataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MistyRose;
			}
			if (typ == 1)
			{
				this.DataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.AliceBlue;
			}
			if (typ == 2)
			{
				this.DataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Coral;
			}
		}

		public static string Get_date_month(DateTime date_in)
		{
            string mon = date_in.Month.ToString("00");
            return date_in.Year + "." + mon;
		}

		private void Type_proc_SelectedIndexChanged(System.Object sender, System.EventArgs e)
		{
			if (type_proc.Text == "аннуитетная")
			{
				this.type_r_stavka.Enabled = true;
                this.sum_plat.Enabled = true;
                this.groupbox_rasrochka.Visible = false;
            }
			else if (type_proc.Text == "классика")
			{
				this.type_r_stavka.Enabled = true;
                this.sum_plat.Enabled = true;
                this.groupbox_rasrochka.Visible = false;
            }
            else
            {
                this.type_r_stavka.Enabled = false;
                this.sum_plat.Enabled = true;
                this.type_r_stavka.Text = "факт/факт";
                this.groupbox_rasrochka.Visible = true;
            }
        }

#endregion

        // Загрузка
		private void Kred_calculator_Load(object sender, System.EventArgs e)
		{
            Poisk_ini_files();
            if (this.type_rasch.Items.Count == 0)
            {
                MessageBox.Show("Не найдены файлы ini !!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }

            this.type_rasch.SelectedIndex = 0;
			Start_load_param();
		}

        // Загрузка начальных параметров
		public void Start_load_param()
		{
			Load_radio();
			Read_ini_file_kred_calc();
			type_r_stavka.SelectedIndex = 2;
			Paint_table();
			F_proc_perv();

			if (Prov_numeric(this.summa_ekv.Text) != 0 && Prov_numeric(this.summa.Text) == 0)
			{
				this.summa.Text = null;
				this.kurs.Text = null;
			}
            this.sum_plat.Text = "0";
            this.srok_cred_new.Text = "";
            this.srok_kred_year_new.Text = "";
            this.pereplata.Text = "0";
        }


        // Расчитать
		private void Button1_Click(System.Object sender, System.EventArgs e)
		{
			F_proc_perv();
			Rashet_formula(this.type_proc.Text, this.type_r_stavka.Text);
            if (this.DataGridView1.Rows.Count >= 3)
            {
                //this.sum_na_ruki.Text = (double.Parse(this.sum_na_ruki.Text) + double.Parse(this.DataGridView1.Rows[0].Cells[5].Value.ToString())).ToString("#,0.00");
            }
		}

        // Пересчет первоначального взноса
		public void F_proc_perv()
		{
			if (this.CheckBox1.Checked == false)
			{
                this.perv_vznos.Enabled = true;
                this.proc_perv_vznos.Enabled = false;
                F_proc_perv_vznos();
            }
			else
			{
                this.perv_vznos.Enabled = false;
                this.proc_perv_vznos.Enabled = true;
                F_proc_perv_vznos_proc();
            }
		}

        // Пересчет первоначального взноса - сумма
        public void F_proc_perv_vznos()
		{
            double p = Prov_numeric(this.perv_vznos.Text);

            double s;
            if (Prov_numeric(this.summa_ekv.Text) != 0 && Prov_numeric(this.summa.Text) == 0)
            {
                s = Prov_numeric(this.summa_ekv.Text);
            }
            else
            {
                s = Prov_numeric(this.summa.Text) * Prov_numeric(this.kurs.Text);
                this.summa_ekv.Text = s.ToString("#,0.00");
            }

            if (s < 0)
			{
				this.sum_kred.Text = "#Error#";
				return;
			}

			if (p < 0 && !string.IsNullOrEmpty(this.perv_vznos.Text))
			{
				this.sum_kred.Text = "#Error#";
				return;
			}

			if (s < p)
			{
				MessageBox.Show("Первоначальный взнос не может быть больше стоимости квартиры !!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				p = 0;
				this.perv_vznos.Text = "0";
			}

			if (s != 0)
			{
				this.proc_perv_vznos.Text = (Round((p / s) * 100, 2)).ToString("#,0.00");
			}
			else
			{
				this.proc_perv_vznos.Text = "";
			}

			this.sum_kred.Text = (Round(s - p, 2)).ToString("#,0.00");
			Dop_plat();
		}

        // Пересчет первоначального взноса - процент
        public void F_proc_perv_vznos_proc()
		{
            double p = Prov_numeric(this.proc_perv_vznos.Text);

            double s;
            if (Prov_numeric(this.summa_ekv.Text) != 0 && Prov_numeric(this.summa.Text) == 0)
            {
                s = Prov_numeric(this.summa_ekv.Text);
            }
            else
            {
                s = Prov_numeric(this.summa.Text) * Prov_numeric(this.kurs.Text);
                this.summa_ekv.Text = s.ToString("#,0.00");
            }

            if (s < 0)
			{
				this.sum_kred.Text = "#Error#";
				return;
			}

			if (p < 0 && !string.IsNullOrEmpty(this.proc_perv_vznos.Text))
			{
				this.sum_kred.Text = "#Error#";
				return;
			}

			if (s < p)
			{
				MessageBox.Show("Первонаачльный взнос не может быть больше стоимости квартиры !!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				p = 0;
				this.perv_vznos.Text = "0";
			}

			if (s != 0)
			{
				this.perv_vznos.Text = (Round((p * s) / 100, 2)).ToString("#,0.00");
			}
			else
			{
				this.perv_vznos.Text = "";
			}

			this.sum_kred.Text = (Round(s - double.Parse(this.perv_vznos.Text), 2)).ToString("#,0.00");
			Dop_plat();
		}

        // Пересчитать параметры
        private void Button2_Click(System.Object sender, System.EventArgs e)
		{			
			F_proc_perv();
		}

        // Изменение срока кредита в месяцах
		private void Srok_kred_TextChanged(System.Object sender, System.EventArgs e)
		{
			this.srok_kred_year.Text = (Round(Prov_numeric(this.srok_kred.Text) / 12, 2)).ToString("#,0.00");
		}

        // INI файл
        private void Button3_Click(System.Object sender, System.EventArgs e)
		{
            Process p = new();
            p.StartInfo.UseShellExecute = true;
            p.StartInfo.FileName = file_path_ini_mas[this.type_rasch.SelectedIndex];
            p.Start();
		}

        // Перезаполнить
        private void Button4_Click(System.Object sender, System.EventArgs e)
		{
			Start_load_param();
		}

        // Смена типа расчета
		private void Type_rasch_SelectedIndexChanged(System.Object sender, System.EventArgs e)
		{
			Start_load_param();
		}

        // Экспорт в Excel
		private void ЭкспортВCSVToolStripMenuItem_Click(System.Object sender, System.EventArgs e)
		{
            SaveFileDialog OpenSavefileDialog = new()
            {
                Filter = "CSV file|*.csv",
                Title = "Save an CSV File"
            };
            OpenSavefileDialog.ShowDialog();

            if (OpenSavefileDialog.FileName != "")
            {
                string filename = OpenSavefileDialog.FileName.ToString();
                try
                {
                    ToCSV(this.DataGridView1, filename);
                }
                catch
                {
                    MessageBox.Show("Ошибка выгрузки файла CSV !!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
		}

        private static void ToCSV(DataGridView DataGridView1, string strFilePath)
        {
            StreamWriter sw = new(strFilePath, false, System.Text.Encoding.Default);
            //headers  
            for (int i = 0; i < DataGridView1.ColumnCount; i++)
            {
                sw.Write(DataGridView1.Columns[i].HeaderText.ToString());
                if (i < DataGridView1.ColumnCount - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            for (int ii = 0; ii < DataGridView1.RowCount; ii++)
            {
                for (int i = 0; i < DataGridView1.ColumnCount; i++)
                {
                    if (!Convert.IsDBNull(DataGridView1[i, ii].Value))
                    {
                        string? value = DataGridView1[i, ii].Value.ToString();
                        if (value != null) { value = value.Replace(" ", null); }
                        if (value != null)
                        {
                            if (value.Contains(';'))
                            {
                                value = String.Format("\"{0}\"", value);
                            }
                        }
                        sw.Write(value);
                    }
                    if (i < DataGridView1.ColumnCount - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
        
        // Смена кода валюты
        private void Curr_code_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.curr_code.Text == "")
            {
                return;
            }                               

            if (this.curr_code.Text == "UAH") {
                this.kurs.Text = "1";
                this.kurs.Enabled = false;
            }
            else {
                this.kurs.Enabled = true;
                if (!string.IsNullOrEmpty(IniFile.IniReadValue("MAIN", "KURS")))
                {
                    this.kurs.Text = IniFile.IniReadValue("MAIN", "KURS");
                }              
                else {                    
                    this.kurs.Text = Get_kurs_nbu_in_site(this.curr_code.Text, this.date_cred.Value);
                }
            }
        }

        // Чтение курса НБУ с сайта
        private string Get_kurs_nbu_in_site(string p_curr_code, DateTime p_date)
        {
            string url = "https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?valcode=" + p_curr_code + "&date=" + p_date.ToString("yyyyMMdd");
            string rezult = "";
            double rezult_dbl;

            int num_curr;
            int num = curr_code_date.IndexOf(url);
            if (num >= 0)
            {
                num_curr = 0;
                foreach (string ccv in curr_code_val)
                {
                    if (num_curr == num) { return ccv.ToString(); }
                    num_curr++;
                }
            }            

            // забираем файл
            try
            {
                XmlTextReader reader = new(url);
                while (reader.Read())
                {
                    if (reader.Name == "rate")
                    {
                        rezult = reader.ReadString().Replace(".",",");
                        rezult_dbl = Round(double.Parse(rezult),3);
                        rezult = rezult_dbl.ToString("0.000");
                    }                                        
                }
            }
            catch
            {
                MessageBox.Show("Ошибка !!! курс с сайта НБУ загрузить не получилось !!!", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                rezult = "1";
            }

            // Добавляем данные в последовательность
            curr_code_date.Add(url);
            curr_code_val.Add(rezult);

            return rezult;
        }

        // Смена даты оформления кредита
        private void Date_cred_ValueChanged(object sender, EventArgs e)
        {
            Curr_code_SelectedIndexChanged(sender, e);

            this.year_0.Text = this.date_cred.Value.Year.ToString();
            this.year_1.Text = this.date_cred.Value.AddYears(1).Year.ToString();
            this.year_2.Text = this.date_cred.Value.AddYears(2).Year.ToString();
            this.year_3.Text = this.date_cred.Value.AddYears(3).Year.ToString();
            this.year_4.Text = this.date_cred.Value.AddYears(4).Year.ToString();
        }

        // Изменение льготного срока кредита в месяцах
        private void Priv_srok_kred_TextChanged(object sender, EventArgs e)
        {
            this.priv_srok_kred_year.Text = (Round(Prov_numeric(this.priv_srok_kred.Text) / 12, 2)).ToString("#,0.00");
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.CheckBox1.Checked == false)
            {
                this.proc_perv_vznos.Enabled = false;
                this.perv_vznos.Enabled = true;
            }
            else
            {
                this.proc_perv_vznos.Enabled = true;
                this.perv_vznos.Enabled = false;
            }
        }
    }
}