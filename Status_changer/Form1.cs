using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using teemtalk;
using System.Globalization;
using System.Diagnostics;
using NLog;
using Excel = Microsoft.Office.Interop.Excel;


namespace Status_changer
{
    
    public partial class Form1 : Form
    {

        private static Logger logger = LogManager.GetCurrentClassLogger(); // Nlog

        public Form1()
        {
            InitializeComponent();
        }



        static teemtalk. Application teemApp;

        public string EventDepot { get; private set; }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {

            

                //поиск файла Excel
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = false;
                ofd.DefaultExt = "*.xls;*.xlsx";
                ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
                ofd.Title = "Выберите документ Excel";
                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string xlFileName = ofd.FileName; //имя нашего Excel файла

                Excel.Application ObjWorkExcel = new Excel.Application(); //создаём приложение Excel
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(xlFileName); //открываем наш файл 
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[3]; //получить 3 лист





                //Login into Mainframe
                //var login = textBox1.Text;
                //var password = textBox2.Text;
                var login = Properties.Settings.Default.loginMF;
                var password = Properties.Settings.Default.pwdMF;
                //var consData = DBContext.GetConsStatus();
                teemApp = new teemtalk.Application();
            
                teemApp.CurrentSession.Name = "Mainframe";

                teemApp.CurrentSession.Network.Protocol = ttNetworkProtocol.ProtocolWinsock;
                teemApp.CurrentSession.Network.Hostname = "mainframe.gb.tntpost.com";
                teemApp.CurrentSession.Network.Telnet.Port = 23;
                teemApp.CurrentSession.Network.Telnet.Name = "IBM-3278-2-E";
                teemApp.CurrentSession.Emulation = ttEmulations.IBM3270Emul;

                teemApp.CurrentSession.Network.Connect();

                teemApp.Visible = Properties.Settings.Default.isVisible;
            

                var host = teemApp.CurrentSession.Host;
                var disp = teemApp.CurrentSession.Display;

                ForAwait(35, 16, "INTERNATIONAL");

                host.Send("SM");
                host.Send("<ENTER>");

                ForAwait(13, 23, "USER ID");
                Thread.Sleep(2000);
                host.Send(login);
                host.Send("<TAB>");
                host.Send(password);
                host.Send("<ENTER>");

                //if (!ForAwait(2, 2, "Command")) goto StartMaimframe;
                ForAwait(2, 2, "Command");
                host.Send("2");
                host.Send("<ENTER>");

                ForAwait(20, 7, "Job Description");
                host.Send("<F12>");
                Thread.Sleep(500);
                if (disp.CursorRow != 2)

                host.Send("YL30");
                logger.Debug("YL30", this.Text); //LOG
                host.Send("<ENTER>");



                // Загрузка определенного excel файла
                //Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
                //  Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"D:\Projects\Status_changer\Nissan SPR project SPS 1.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                //  Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[3]; //получить 3 лист





                var last = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);//1 ячейку                          
                int lastUsedRow = last.Row;

                //Дата, введенная пользователем в DateSform1
                var DateSform = this.DateSform1.Text;
                //DateTime DateS = DateTime.Parse((string) DateSform);
                logger.Debug(DateSform, this.Text); //LOG


                // foreach (DataRow row in consData.Rows) // Старое
                for (int i = 2; i < lastUsedRow; i++)            
                {               
                    // Colnum
                    string colnum = Convert.ToString(i);

                    

                    // DataS эскпорт                          
                    Excel.Range exceldateS = ObjWorkSheet.get_Range("R" + colnum);             
                    object dateS_v = exceldateS.Value2;

                    if (dateS_v == null || dateS_v is string)
                    {
                        continue; //переход к следующей итерации FOR
                    }

                    DateTime dSt = DateTime.FromOADate((double)dateS_v);

                    string dateS = dSt.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("RU-ru"));
                    
                    if (dateS_v is double)
                    {
                        dSt = DateTime.FromOADate((double)dateS_v);
                    }
                    else
                    {
                        DateTime.TryParse((string)dateS_v, out dSt);
                    }


                    // Done
                    string done = "DONE";


                    //Если статус Введенная дата есть в ячейке, то  цикл продолжается, если нет, то перескакивает к следующему i
                    if (DateSform == dateS) 
                    {
                        logger.Debug(colnum, this.Text); //LOG

                        // Status
                        var excelstatus = ObjWorkSheet.get_Range("Q" + colnum, Type.Missing).Value2;
                        string status = excelstatus.ToString();                                  
                    
                        // Экспорт даты доставки dateD                
                        Excel.Range exceldateD = ObjWorkSheet.get_Range("P" + colnum);
                        object dateD_v = exceldateD.Value2;
                                           
                        // Экспорт даты забора dateZ                
                        Excel.Range exceldateZ = ObjWorkSheet.get_Range("O" + colnum);
                        object dateZ_v = exceldateZ.Value2;
                    
                        //Time - Можно по умолчанию вводить "1000"
                        var time = "1000";

                        //Depo - EVENTDEPOT                     
                        var eventdepot = "MOW";

                        // Consigment номер накладной
                        var excelcon = ObjWorkSheet.get_Range("M" + colnum, Type.Missing).Value2;
                        string con_wocheck = excelcon.ToString();
                        string con = con_wocheck.Substring(0, 9);// ограничивание накладной по количеству знаков

                        // количество обрабатываемых накладных за раз (по умолчанию 1)
                        var qty = "1";

                        // Delv zone - по умолчанию "b"
                        var delvz = "B";

                        ForAwait(15, 2, "Consignment Status Entry");
                        Thread.Sleep(600);                        
                        host.Send(status);//Вводим статус
                        logger.Debug(status, this.Text); //LOG
                        Thread.Sleep(600); //костыль
                        if (disp.CursorCol != 28 && disp.CursorCol != 10)
                            host.Send("<TAB>");

                        ForAwaitCol(28);//Вводим дату доставки, если ОК или дату забора, если OF
                        if (status == "OK")
                        {
                            if (dateD_v == null)// Если даты нет, то переходим к вводу статуса и след. строке
                            {
                                host.Send("<F12>");
                                continue; //переход к следующей итерации FOR
                            }

                            DateTime dDt = DateTime.FromOADate((double)dateD_v);

                            string dateD = dDt.ToString("ddMMMyy", CultureInfo.GetCultureInfo("en-us"));

                            if (dateD_v is double)
                            {
                                dDt = DateTime.FromOADate((double)dateD_v);
                            }
                            else
                            {
                                DateTime.TryParse((string)dateD_v, out dDt);
                            }

                            host.Send(dateD);
                            Thread.Sleep(100);
                            logger.Debug(dateD, this.Text);  //LOG                                                        
                            host.Send("<TAB>");
                        }
                        else if(status == "OF")
                        {
                            if (dateZ_v == null)// Если даты нет, то переходим к вводу статуса и след. строке                    
                            {
                                host.Send("<F12>");
                                continue; //переход к следующей итерации FOR
                            }

                            DateTime dZt = DateTime.FromOADate((double)dateZ_v);

                            string dateZ = dZt.ToString("ddMMMyy", CultureInfo.GetCultureInfo("en-us"));

                            if (dateZ_v is double)
                            {
                                dZt = DateTime.FromOADate((double)dateZ_v);
                            }
                            else
                            {
                                DateTime.TryParse((string)dateZ_v, out dZt);
                            }

                            host.Send(dateZ);
                            Thread.Sleep(100);
                            logger.Debug(dateZ, this.Text);  //LOG                                                       
                            host.Send("<TAB>");
                        }

                        ForAwaitCol(46);//Вводим время
                        host.Send(time);
                        Thread.Sleep(100);
                        if (disp.CursorCol != 70 && disp.CursorCol != 46) host.Send("<TAB>");
                                                
                        ForAwaitCol(70);//Вводим депо
                        Thread.Sleep(3000);
                        host.Send("MOW");
                        Thread.Sleep(3000);
                        host.Send("<TAB>");
                        Thread.Sleep(100);

                        ForAwaitCol(13);// Signatory - пропускаем
                        host.Send("<TAB>");

                        ForAwaitCol(57);// REV Date - пропускаем                    
                        host.Send("<TAB>");

                        ForAwaitCol(77);//Rems + Если статус OF, то делаем и вводим коммент = статусу OF
                        if (status == "OF")
                        {
                            host.Send("<F4>");
                            ForAwait(5, 5, "Seq Remarks");                            
                            host.Send(status);
                            Thread.Sleep(500);
                            host.Send("<ENTER>");

                            ForAwaitCol(9); // вторая строка seq remarks
                            host.Send("<F12>");

                            ForAwaitCol(18);// mode: add - пропускаем
                            host.Send("<F12>");//возвращаемся в общее меню на позицию REMS+ COL(77)
                            ForAwait(15, 2, "Consignment Status Entry");// проверяем                    
                        }                               
                        host.Send("<TAB>");

                        ForAwaitCol(12);//Runsheet - пропускаем
                        host.Send("<TAB>");

                        ForAwaitCol(33);//Round no - пропускаем
                        host.Send("<TAB>");

                        ForAwaitCol(54);// Delv zone -  по умолчанию "b"
                        host.Send(delvz); 
                        host.Send("<TAB>");

                        ForAwaitCol(73);// Delv area - пропускаем
                        host.Send("<TAB>");

                        ForAwaitCol(24);//No of status Entries = 1
                        host.Send(qty);
                        host.Send("<ENTER>");
                        ForAwait(1, 10, "01");

                        host.Send(con);  // Con number        
                        logger.Debug(con, this.Text); //LOG
                        ForAwaitCol(26);//Позиция после ввода 9 символов номера накладной    
                        host.Send("<TAB>");

                        ForAwaitCol(37);// Статус (повторный вывод) - пропускаем
                        host.Send("<TAB>");

                        ForAwaitCol(48);// Time - пропускаем
                        host.Send("<TAB>");

                        ForAwaitCol(58);// Solved - пропускаем
                        host.Send("<TAB>");

                        ForAwaitCol(64);// Rev date (повторный вывод) - пропускаем
                        host.Send("<TAB>");

                        ForAwaitCol(17); // Signatory Если статус OK = OK, если OF = ""
                        if (status == "OK")
                        {
                            host.Send(status);
                        }
                        else
                        {
                            host.Send("");
                        }

                        host.Send("<ENTER>");//концовка и переход обратно к вводу статуса
                        host.Send("<F12>");
                        host.Send("<ENTER>");
                        Thread.Sleep(2500);

                    

                            if (disp.ScreenData[15, 2, 9] == "Duplicate")
                            {
                                var checkDepo = "";
                                short j = 1;
                                do
                                {
                                    short col = (Int16)(9 + j);
                                    checkDepo = disp.ScreenData[54, col, 3];
                                    if (checkDepo == "MW3" || checkDepo == "MW5" || checkDepo == "MW7" || checkDepo == "MOW"

                                        || checkDepo == "LED"
                                        || checkDepo == "KG7"
                                        || checkDepo == "GOJ"
                                        || checkDepo == "KUF"
                                        || checkDepo == "KZ7"
                                        || checkDepo == "RO8"
                                        || checkDepo == "KR4"
                                        || checkDepo == "SVX"
                                        || checkDepo == "IK3"
                                        || checkDepo == "OVB"
                                        || checkDepo == "KH6"
                                        || checkDepo == "VK3"
                                        || checkDepo == "AB7"
                                        || checkDepo == "AC8"
                                        || checkDepo == "AK7"
                                        || checkDepo == "AP6"
                                        || checkDepo == "AV8"
                                        || checkDepo == "BA8"
                                        || checkDepo == "BB8"
                                        || checkDepo == "BG8"
                                        || checkDepo == "BU8"
                                        || checkDepo == "BY5"
                                        || checkDepo == "CB2"
                                        || checkDepo == "CT6"
                                        || checkDepo == "EL6"
                                        || checkDepo == "IV6"
                                        || checkDepo == "IZ8"
                                        || checkDepo == "JA5"
                                        || checkDepo == "KE5"
                                        || checkDepo == "KG5"
                                        || checkDepo == "KI4"
                                        || checkDepo == "KJ4"
                                        || checkDepo == "KM7"
                                        || checkDepo == "KN6"
                                        || checkDepo == "KU3"
                                        || checkDepo == "KU8"
                                        || checkDepo == "LI5"
                                        || checkDepo == "MK5"
                                        || checkDepo == "MU5"
                                        || checkDepo == "MV7"
                                        || checkDepo == "NC8"
                                        || checkDepo == "NH2"
                                        || checkDepo == "NV6"
                                        || checkDepo == "NZ8"
                                        || checkDepo == "OM4"
                                        || checkDepo == "OR7"
                                        || checkDepo == "OR8"
                                        || checkDepo == "PK7"
                                        || checkDepo == "PK9"
                                        || checkDepo == "PS9"
                                        || checkDepo == "PV3"
                                        || checkDepo == "PZ6"
                                        || checkDepo == "RC6"
                                        || checkDepo == "RT4"
                                        || checkDepo == "RT6"
                                        || checkDepo == "RY2"
                                        || checkDepo == "SH5"
                                        || checkDepo == "SH6"
                                        || checkDepo == "SK9"
                                        || checkDepo == "SM2"
                                        || checkDepo == "SP5"
                                        || checkDepo == "SQ4"
                                        || checkDepo == "SR7"
                                        || checkDepo == "SU8"
                                        || checkDepo == "SY5"
                                        || checkDepo == "TB3"
                                        || checkDepo == "TO8"
                                        || checkDepo == "TU3"
                                        || checkDepo == "TV6"
                                        || checkDepo == "UF5"
                                        || checkDepo == "UK4"
                                        || checkDepo == "UL9"
                                        || checkDepo == "UU3"
                                        || checkDepo == "UV4"
                                        || checkDepo == "VL4"
                                        || checkDepo == "VL5"
                                        || checkDepo == "VN4"
                                        || checkDepo == "VO4"
                                        || checkDepo == "VO6"
                                        || checkDepo == "VO8"
                                        || checkDepo == "VY6"
                                        || checkDepo == "VY8"
                                        || checkDepo == "XS7"
                                        || checkDepo == "ZP8"

                                        )
                                    {
                                        host.Send(j.ToString());
                                        host.Send("<ENTER>");
                                        //break;
                                    }
                                   j++;
                                } while (checkDepo.Trim() != "");
                                host.Send("1");
                                host.Send("<ENTER>");
                                Thread.Sleep(2000);

                            }


                        ForAwait(15, 2, "Consignment Status Entry");
                        // DBContext.ChangeRecordStatus(id); 

                         // Запись в ячейку даты внесения статуса отметки DONE
                        //ObjWorkSheet.Cells[18, i] = done;
                        //ObjWorkExcel.Interactive = false;
                        //ObjWorkBook.Save();
                        //ObjWorkExcel.Interactive = true;
                        logger.Debug(done, this.Text);  //LOG
                    }
                    else
                    {
                        continue; //переход к следующей итерации FOR
                    }                               
                
                }


                // Закрываем TeemTalk
                teemApp.Close();
                foreach (Process proc in Process.GetProcessesByName("teem2k"))
                {
                    proc.Kill();
                }
                //teemApp.Application.Close();
                Thread.Sleep(1000);
                //host.Send("<ENTER>");



                //Закрываем Excel
                ObjWorkBook.Close(true);
                ObjWorkExcel.Quit();
                foreach (Process proc in Process.GetProcessesByName("excel"))
                {
                    proc.Kill();
                }

                MessageBox.Show("Данные внесены", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;

            }
            catch (Exception ex)
            {
                // Вывод сообщения об ошибке
                logger.Debug(ex.ToString());
            }




        }




        static bool ForAwait(short col, short row, string keyword)
        {
            byte count = 0;
            
                do
                {
                    count++;
                    
                    if (count > 70)
                    {
                        teemApp.CurrentSession.Network.Close();
                        Thread.Sleep(1000);
                        teemApp.Close();

                        System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                        foreach (System.Diagnostics.Process p in process)
                        {
                            if (!string.IsNullOrEmpty(p.ProcessName))
                            {
                                try
                                {
                                    p.Kill();
                                }

                                catch (Exception ex)
                                {
                                    // Вывод сообщения об ошибке
                                    logger.Debug(ex.ToString());
                                }
                            }
                        }

                        return false;
                    }

                    Thread.Sleep(100);

                } while ((teemApp.CurrentSession.Display.ScreenData[col, row, (short)keyword.Length] != keyword));
            return true;
        }

        static bool ForAwaitRow(short keyword)
        {
            byte count = 0;

            do
            {
                count++;

                if (count > 70)
                {
                    teemApp.CurrentSession.Network.Close();
                    Thread.Sleep(1000);
                    teemApp.Close();

                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch (Exception ex)
                            {
                                // Вывод сообщения об ошибке
                                logger.Debug(ex.ToString());
                            }
                        }
                    }

                    return false;
                }

                Thread.Sleep(100);

            } while ((teemApp.CurrentSession.Display.CursorRow != keyword));
            return true;
        }
        static bool ForAwaitCol(short keyword)
        {
            byte count = 0;

            do
            {
                count++;

                if (count > 70)
                {
                    teemApp.CurrentSession.Network.Close();
                    Thread.Sleep(1000);
                    teemApp.Close();

                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch (Exception ex)
                            {
                                // Вывод сообщения об ошибке
                                logger.Debug(ex.ToString());
                            }
                        }
                    }

                    return false;
                }

                Thread.Sleep(100);

            } while ((teemApp.CurrentSession.Display.CursorCol != keyword));
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
