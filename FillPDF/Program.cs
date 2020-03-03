using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Security;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace FillPDF
{
    class Program
    {
        enum PublisherType { RegularPioneer, NonBaptized, Baptized, Undefined }

        static void Main()
        {
            var currentMonthStr = DateTime.Now.ToString("MM");
            if (!int.TryParse(currentMonthStr, out int currentMonth))
            {
                Console.WriteLine($"Erro ao converter mês corrente para inteiro. Mês: {currentMonthStr}");
                goto Finish;
            }

            var yearStr = DateTime.Now.ToString("yyyy");
            if (!int.TryParse(yearStr, out int serviceYear))
            {
                Console.WriteLine($"Erro ao converter ano corrente para inteiro. Ano: {yearStr}");
                goto Finish;
            }

            if (currentMonth >= 9) serviceYear = serviceYear + 1;

            var pathRelatorios = Path.Combine(Directory.GetCurrentDirectory(), "Relatorios");

            if (!Directory.Exists(pathRelatorios))
            {
                Directory.CreateDirectory(pathRelatorios);
            }

            var filename = Directory.GetFiles(pathRelatorios).FirstOrDefault(file => file.EndsWith(".xml"));

            if (filename == null)
            {
                Console.WriteLine("Não existe nenhum ficheiro XML na pasta Relatorios!");
                goto Finish;
            }

            var currentDirectory = Directory.GetCurrentDirectory();
            var doc = Path.Combine(currentDirectory, filename);

            XElement root = XElement.Load(doc);

            IEnumerable<XElement> pubs = from el in root.Elements("Active").Elements("Pub")
                                         select el;

            if (!int.TryParse((string)root.Element("Count"), out int totalPub))
            {
                Console.WriteLine("Número de Publicadores Activos no xml não é um inteiro válido.");
                goto Finish;
            }

            Console.WriteLine("Publicadores Activos: " + totalPub);

            string folderName = "Export-" + DateTime.Now.ToString("dd-MM-yyyy_hh-mm-ss");
            Directory.CreateDirectory(folderName);

            int index = 1;

            foreach (XElement el in pubs)
            {
                FileStream file;
                try
                {
                    file = new FileStream("S-21-TPO.pdf", FileMode.Open);
                } catch(FileNotFoundException)
                {
                    Console.WriteLine("Coloque o PDF S-21 na pasta do executável, com o nome: S-21-TPO.pdf");
                    goto Finish;
                }
                //Load the PDF document
                PdfLoadedDocument loadedDocument = new PdfLoadedDocument(file);

                ////Gets the first page of the document
                PdfLoadedPage page = loadedDocument.Pages[0] as PdfLoadedPage;

                ////Get the loaded form
                PdfLoadedForm form = loadedDocument.Form;

                // Print Fields' Names

                //foreach (PdfLoadedField field in form.Fields)
                //{
                //    Console.WriteLine(field.Name);
                //}

                var pubName = (string)el.Element("fname") + " " + (string)el.Element("lname");

                Console.WriteLine($"A preencher: {pubName} ({index++}/{totalPub})");

                var publisherType = FillForm(form, el, serviceYear.ToString(),2);
                FillForm(form, el, (serviceYear-1).ToString(),1);

                string subfolder = "Outros";

                switch (publisherType)
                {
                    case PublisherType.RegularPioneer:
                        subfolder = "Pioneiros Regulares";
                        break;
                    case PublisherType.NonBaptized:
                        subfolder = "Publicadores Não Batizados";
                        break;
                    case PublisherType.Baptized:
                        subfolder = $"Grupo {(string)el.Element("group")}";
                        break;
                }

                SaveFile(loadedDocument, folderName, subfolder, pubName);
            }

            Console.WriteLine("Sucesso!");

        Finish:
            Console.WriteLine("Click em Enter para terminar...");
            Console.ReadLine();
        }

        public static void SaveFile(PdfLoadedDocument loadedDocument, string folderName, string subFolder, string pubName)
        {
            Directory.CreateDirectory(Path.Combine(Directory.GetCurrentDirectory(), folderName, subFolder));

            var fs = File.Create($"{folderName}\\{subFolder}\\{pubName}.pdf");

            // Save the document
            loadedDocument.Save(fs);
            // Close the document
            loadedDocument.Close(true);
            // This will open the PDF file so, the result will be seen in default PDF viewer
            // Process.Start("Form.pdf");
        }

        static PublisherType FillForm(PdfLoadedForm form, XElement el, string year, int tableIndex = 1)
        {
            bool isPionner = false;

            PublisherType publisherType = PublisherType.Undefined;

            // Name
            (form.Fields["Name"] as PdfLoadedTextBoxField).Text = (string)el.Element("fname") + " " + (string)el.Element("lname");

            // Gender
            if ((string)el.Element("gender") == "Male")
            {
                (form.Fields["Check Box1"] as PdfLoadedCheckBoxField).Checked = true;
            }
            else
            {
                (form.Fields["Check Box2"] as PdfLoadedCheckBoxField).Checked = true;
            }

            // Outra ovelha
            (form.Fields["Check Box3"] as PdfLoadedCheckBoxField).Checked = true;

            // <svt>Elder</svt> ,  <svt>MinSvt</svt>, 
            if ((string)el.Element("svt") == "Elder")
            {
                (form.Fields["Check Box5"] as PdfLoadedCheckBoxField).Checked = true;
            }

            if ((string)el.Element("svt") == "MinSvt")
            {
                (form.Fields["Check Box6"] as PdfLoadedCheckBoxField).Checked = true;
            }

            // <birdate>1961-06-14</birdate>
            (form.Fields["Date of birth"] as PdfLoadedTextBoxField).Text = (string)el.Element("birdate");

            //< bapdate > 2003 - 07 - 19 </ bapdate >
            (form.Fields["Date immersed"] as PdfLoadedTextBoxField).Text = (string)el.Element("bapdate");
            if((string)el.Element("bapdate") == String.Empty)
            {
                publisherType = PublisherType.NonBaptized;
            }

            if(tableIndex == 1) (form.Fields["Service Year"] as PdfLoadedTextBoxField).Text = year;
            if(tableIndex == 2) (form.Fields["Service Year_2"] as PdfLoadedTextBoxField).Text = year;

            int totalPlace = 0, totalVideos = 0, totalHours = 0, totalRV = 0, totalStudies = 0, monthsWithActivity = 0;

            string yearBackup = year;

            for (int i = 1; i <= 12; i++)
            {
                if (i <= 4)
                {
                    year = GetPastYear(yearBackup);
                }
                else
                {
                    year = yearBackup;
                }

                var matches = el.Elements(GetMonth3Letters(i));
                var e = matches.FirstOrDefault(x => x.Attribute("Year").Value.Equals(year));
                if (e != null)
                {
                    var value = (string)e.Element("Plcmts");
                    (form.Fields[$"{tableIndex}-Place_{i}"] as PdfLoadedTextBoxField).Text = value;
                    int.TryParse(value, out int v);
                    totalPlace += v;

                    value = (string)e.Element("Videos");
                    (form.Fields[$"{tableIndex}-Video_{i}"] as PdfLoadedTextBoxField).Text = value;
                    int.TryParse(value, out v);
                    totalVideos += v;

                    value = (string)e.Element("Hours");
                    (form.Fields[$"{tableIndex}-Hours_{i}"] as PdfLoadedTextBoxField).Text = value;
                    int.TryParse(value, out v);
                    totalHours += v;

                    value = (string)e.Element("R.V.s");
                    (form.Fields[$"{tableIndex}-RV_{i}"] as PdfLoadedTextBoxField).Text = value;
                    int.TryParse(value, out v);
                    totalRV += v;

                    value = (string)e.Element("BiSt.");
                    (form.Fields[$"{tableIndex}-Studies_{i}"] as PdfLoadedTextBoxField).Text = value;
                    int.TryParse(value, out v);
                    totalStudies += v;

                    value = (string)e.Element("Remark");
                    var text = (string)e.Element("Pio") == "Aux" ? $"Pioneir{(IsMale(el) ? "o" : "a")} Auxiliar. {value}" : value;
                    if (tableIndex == 1) (form.Fields[$"Remarks{GetMonth(i)}"] as PdfLoadedTextBoxField).Text = text;
                    else (form.Fields[$"Remarks{GetMonth(i)}_2"] as PdfLoadedTextBoxField).Text = text;

                    isPionner = (string)e.Element("Pio") == "Reg";
                    (form.Fields["Check Box7"] as PdfLoadedCheckBoxField).Checked = isPionner;
                    if(isPionner)
                    {
                        publisherType = PublisherType.RegularPioneer;
                    }

                    if (!(form.Fields[$"{tableIndex}-Hours_{i}"] as PdfLoadedTextBoxField).Text.Equals(""))
                    {
                        monthsWithActivity++;
                    }
                }

            }

            (form.Fields[$"{tableIndex}-Place_Total"] as PdfLoadedTextBoxField).Text = totalPlace.ToString();
            (form.Fields[$"{tableIndex}-Place_Average"] as PdfLoadedTextBoxField).Text = GetAvg(totalPlace, monthsWithActivity);

            (form.Fields[$"{tableIndex}-Video_Total"] as PdfLoadedTextBoxField).Text = totalVideos.ToString();
            (form.Fields[$"{tableIndex}-Video_Average"] as PdfLoadedTextBoxField).Text = GetAvg(totalVideos, monthsWithActivity);

            (form.Fields[$"{tableIndex}-Hours_Total"] as PdfLoadedTextBoxField).Text = totalHours.ToString();
            (form.Fields[$"{tableIndex}-Hours_Average"] as PdfLoadedTextBoxField).Text = GetAvg(totalHours, monthsWithActivity);

            (form.Fields[$"{tableIndex}-RV_Total"] as PdfLoadedTextBoxField).Text = totalRV.ToString();
            (form.Fields[$"{tableIndex}-RV_Average"] as PdfLoadedTextBoxField).Text = GetAvg(totalRV, monthsWithActivity);

            (form.Fields[$"{tableIndex}-Studies_Total"] as PdfLoadedTextBoxField).Text = totalStudies.ToString();
            (form.Fields[$"{tableIndex}-Studies_Average"] as PdfLoadedTextBoxField).Text = GetAvg(totalStudies, monthsWithActivity);

            return publisherType == PublisherType.Undefined ? PublisherType.Baptized : publisherType;
        }


        static bool IsMale(XElement el) => (string)el.Element("gender") == "Male";

        static string GetAvg(int total, int months) => ((double)total / (double)months).ToString();

        static string GetMonth(int month) => new DateTime(2019, month, 1).AddMonths(8).ToString("MMMM", CultureInfo.CreateSpecificCulture("en"));

        static string GetMonth3Letters(int month) => new DateTime(2019, month, 1).AddMonths(8).ToString("MMM", CultureInfo.CreateSpecificCulture("en"));

        static string GetPastYear(string currentYear)
        {
            int.TryParse(currentYear, out int pastYear);
            return (pastYear-1).ToString();
        }
    }
}
