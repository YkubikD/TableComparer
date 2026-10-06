using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using ExcelDataReader;

namespace TableComparer
{
    public class ProductItem
    {
        public string OrderRow { get; set; } = "";
        public string TtnRow { get; set; } = "";
        public string Article { get; set; } = "";
        public string FullName { get; set; } = "";
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal TotalSum { get; set; }
        public bool IsMatched { get; set; } = false;
    }

    class Program
    {
        // Облачный словарь: Ключ = Код 1С (Заказ), Значение = Текст для поиска (ТТН)
        static Dictionary<string, string> GoogleCloudDictionary = new Dictionary<string, string>();

        [STAThread]
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.Title = "Сверка ТТН и Заказа (Облачная база)";
            Console.WriteLine("=== ЗАПУСК ПРОГРАММЫ СВЕРКИ ===");

            // Ваша новая прямая ссылка на опубликованный Лист 4 в формате CSV
            string googleUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSKW70yWRuYKcwWBXWR38adHM5FtAgZC7VWx08-kGDZb46lgNXMUCsLhrGZroA8pV3Beseo0DukwgQ6/pub?gid=1609197053&single=true&output=csv";

            Console.WriteLine("\n[1/3] Загрузка облачной базы артикулов...");
            if (!LoadGoogleSpreadsheet(googleUrl))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Критическая ошибка: База не загружена. Работа остановлена.");
                Console.ResetColor();
                Console.ReadLine();
                return;
            }

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("\nВыберите файл Накладной (ТТН)...");
            Console.ResetColor();
            string pathFile1 = SelectExcelFile();
            if (string.IsNullOrEmpty(pathFile1)) return;

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Выберите файл Заказа от покупателя...");
            Console.ResetColor();
            string pathFile2 = SelectExcelFile();
            if (string.IsNullOrEmpty(pathFile2)) return;

            Console.WriteLine("\n[2/3] Считывание файлов Excel...");
            List<ProductItem> ttnList = ReadTtnDocument(pathFile1, startRow: 40, nameIdx: 0, qtyIdx: 17, priceIdx: 20);
            List<ProductItem> orderList = ReadOrderDocument(pathFile2, startRow: 25, artIdx: 69, nameIdx: 26, qtyIdx: 101, priceIdx: 119);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Успешно: ТТН = {ttnList.Count} поз. | Заказ = {orderList.Count} поз.");
            Console.ResetColor();

            // === ДИАГНОСТИКА ===
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine("\n--- ТТН (первые 5 позиций) ---");
            foreach (var t in ttnList.Take(5))
                Console.WriteLine($"  стр.{t.TtnRow}: qty={t.Quantity} price={t.Price} | {t.FullName}");

            Console.WriteLine("\n--- Заказ (первые 5 позиций) ---");
            foreach (var o in orderList.Take(5))
                Console.WriteLine($"  стр.{o.OrderRow}: art={o.Article} qty={o.Quantity} price={o.Price} | {o.FullName}");
            Console.ResetColor();
            // === /ДИАГНОСТИКА ===

            Console.WriteLine("\n[3/3] Выполнение автоматической сверки...");
            List<string> reportLines = ExecuteComparison(ttnList, orderList);

            Console.WriteLine("\n=== РЕЗУЛЬТАТЫ СВЕРКИ ===");
            foreach (var line in reportLines)
            {
                if (line.StartsWith("[OK]")) Console.ForegroundColor = ConsoleColor.Green;
                else if (line.StartsWith("[-]")) Console.ForegroundColor = ConsoleColor.Red;
                else if (line.StartsWith("[!]")) Console.ForegroundColor = ConsoleColor.Yellow;
                else if (line.StartsWith("[+]")) Console.ForegroundColor = ConsoleColor.Magenta;
                else if (line.StartsWith("===")) Console.ForegroundColor = ConsoleColor.Cyan;

                Console.WriteLine(line);
                Console.ResetColor();
            }
            Console.WriteLine("=========================\n");

            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("Готово. Нажмите Enter для выхода...");
            Console.ReadLine();
        }

        // =================================================================================
        // ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ДЛЯ НАДЁЖНОГО ПАРСИНГА ЧИСЕЛ ИЗ EXCEL
        // =================================================================================
        static int ReadInt(object obj)
        {
            if (obj == null) return 0;
            if (obj is double d) return (int)Math.Round(d);
            if (obj is decimal m) return (int)Math.Round(m);
            if (obj is int i) return i;
            if (obj is long l) return (int)l;
            int.TryParse(obj.ToString(), out int v);
            return v;
        }

        static decimal ReadDecimal(object obj)
        {
            if (obj == null) return 0m;
            if (obj is double d) return (decimal)d;
            if (obj is decimal m) return m;
            if (obj is int i) return i;
            if (obj is long l) return l;
            var s = obj.ToString().Replace(',', '.');
            decimal.TryParse(s, System.Globalization.NumberStyles.Any,
                             System.Globalization.CultureInfo.InvariantCulture, out var v);
            return v;
        }

        // =================================================================================
        // ЗАГРУЗКА ОБЛАЧНОЙ БАЗЫ
        // =================================================================================
        static bool LoadGoogleSpreadsheet(string url)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var stream = client.GetStreamAsync(url).Result;

                    using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(stream, Encoding.UTF8))
                    {
                        parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                        parser.SetDelimiters(",");
                        parser.HasFieldsEnclosedInQuotes = true;

                        bool isHeader = true;

                        while (!parser.EndOfData)
                        {
                            string[] parts = parser.ReadFields();
                            if (isHeader) { isHeader = false; continue; }

                            if (parts == null || parts.Length < 2) continue;

                            string fullName = parts[0]?.Trim() ?? "";
                            string article = parts[1]?.Trim() ?? "";

                            if (string.IsNullOrEmpty(article) || article.ToLower().Contains("штамп"))
                                break;

                            if (string.IsNullOrEmpty(fullName))
                            {
                                fullName = "Модель без имени (" + article + ")";
                            }

                            if (article.ToLower().Contains("артикул"))
                                continue;

                            string cleanArtKey = article.Replace(" ", "").Replace("-", "").ToUpper();

                            if (!GoogleCloudDictionary.ContainsKey(cleanArtKey))
                            {
                                GoogleCloudDictionary.Add(cleanArtKey, fullName);
                            }
                        }
                    }
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Успешно. Облачная база загружена: {DateTime.Now:dd.MM.yyyy HH:mm}");
                Console.WriteLine($"Активных позиций в памяти: {GoogleCloudDictionary.Count} шт.");
                Console.ResetColor();
                return true;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Ошибка сети или парсера при загрузке базы: {ex.Message}");
                Console.ResetColor();
                return false;
            }
        }

        static string SelectExcelFile()
        {
            using (var ofd = new System.Windows.Forms.OpenFileDialog())
            {
                ofd.Filter = "Файлы Excel (*.xlsx;*.xls)|*.xlsx;*.xls";
                ofd.Title = "Выберите документ";
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Console.WriteLine($"Выбран: {Path.GetFileName(ofd.FileName)}");
                    return ofd.FileName;
                }
            }
            return null;
        }

        // =================================================================================
        // МЕТОД СВЕРКИ (с проверкой КОЛИЧЕСТВА, ЦЕНЫ, СУММЫ + вывод [OK] + СВОДКА)
        // =================================================================================
        static List<string> ExecuteComparison(List<ProductItem> ttnList, List<ProductItem> orderList)
        {
            List<string> report = new List<string>();

            foreach (var orderItem in orderList)
            {
                string cleanOrderArt = orderItem.Article.Replace(" ", "").Replace("-", "").ToUpper();

                if (!GoogleCloudDictionary.ContainsKey(cleanOrderArt))
                {
                    report.Add($"[-] Товар из Заказа [{orderItem.Article}] ({orderItem.FullName}) ОТСУТСТВУЕТ в Гугл-таблице! Внесите его в базу.");
                    continue;
                }

                string cloudTtnNameTarget = GoogleCloudDictionary[cleanOrderArt].ToUpper().Replace(" ", "").Replace("-", "");

                string orderSide = "БЕЗ_МАРКИРОВКИ";
                if (cleanOrderArt.EndsWith("L") || orderItem.FullName.ToUpper().Contains("ЛЕВ") || orderItem.FullName.ToUpper().Contains("СЛЕВА") || orderItem.FullName.ToUpper().Contains("ЧАШАЛЕВАЯ")) orderSide = "ЛЕВАЯ";
                if (cleanOrderArt.EndsWith("R") || orderItem.FullName.ToUpper().Contains("ПРАВ") || orderItem.FullName.ToUpper().Contains("СПРАВА") || orderItem.FullName.ToUpper().Contains("ЧАШАПРАВАЯ")) orderSide = "ПРАВАЯ";

                ProductItem ttnMatch = ttnList.Find(t => {
                    if (t.IsMatched) return false;

                    string cleanTtnName = t.FullName.ToUpper().Replace(" ", "").Replace("-", "");

                    if (!cleanTtnName.Contains(cloudTtnNameTarget)) return false;

                    if (cloudTtnNameTarget.Contains("ЛЕВ") || cloudTtnNameTarget.Contains("ПРАВ")) return true;

                    string ttnSide = "БЕЗ_МАРКИРОВКИ";
                    if (cleanTtnName.Contains("ЛЕВ") || cleanTtnName.Contains("СЛЕВА")) ttnSide = "ЛЕВАЯ";
                    if (cleanTtnName.Contains("ПРАВ") || cleanTtnName.Contains("СПРАВА")) ttnSide = "ПРАВАЯ";

                    return orderSide == ttnSide;
                });

                if (ttnMatch != null)
                {
                    orderItem.IsMatched = true; ttnMatch.IsMatched = true;

                    bool qtyMismatch = orderItem.Quantity != ttnMatch.Quantity;
                    bool priceMismatch = Math.Round(orderItem.Price, 2) != Math.Round(ttnMatch.Price, 2);
                    bool sumMismatch = Math.Round(orderItem.TotalSum, 2) != Math.Round(ttnMatch.TotalSum, 2);

                    if (qtyMismatch || priceMismatch || sumMismatch)
                    {
                        string err = $"[!] Расхождение (Код 1С: {orderItem.Article}):\n" +
                                     $"    Заказ: {orderItem.FullName}\n" +
                                     $"    ТТН:   {ttnMatch.FullName}\n";

                        if (qtyMismatch)
                            err += $"    [-] КОЛИЧЕСТВО: Заказ = {orderItem.Quantity} | ТТН = {ttnMatch.Quantity}\n";

                        if (priceMismatch)
                            err += $"    [-] ЦЕНА ЗА ШТ: Заказ = {orderItem.Price:F2} | ТТН = {ttnMatch.Price:F2} " +
                                   $"(разница {orderItem.Price - ttnMatch.Price:F2})\n";

                        if (sumMismatch)
                            err += $"    [-] СУММА: Заказ = {orderItem.TotalSum:F2} | ТТН = {ttnMatch.TotalSum:F2}\n";

                        report.Add(err);
                    }
                    else
                    {
                        // ВСЁ СОШЛОСЬ
                        report.Add($"[OK] {orderItem.Article} | {orderItem.FullName} | " +
                                   $"кол-во = {orderItem.Quantity} | цена = {orderItem.Price:F2} | сумма = {orderItem.TotalSum:F2}");
                    }
                }
                else
                {
                    report.Add($"[-] Товар из Заказа [{orderItem.Article}] ({orderItem.FullName}) НЕ НАЙДЕН в ТТН отгрузки! " +
                               $"Не хватает: {orderItem.Quantity} шт. | цена = {orderItem.Price:F2} | сумма = {orderItem.TotalSum:F2}");
                }
            }

            // Лишние позиции в ТТН
            foreach (var ttn in ttnList)
            {
                if (!ttn.IsMatched)
                {
                    report.Add($"[+] Лишний товар в ТТН! Позиция [{ttn.FullName}] отсутствует в Заказе клиента.");
                }
            }

            // === СВОДКА ===
            int okCount = report.Count(r => r.StartsWith("[OK]"));
            int mismatchCount = report.Count(r => r.StartsWith("[!]"));
            int notFoundCount = report.Count(r => r.StartsWith("[-] Товар из Заказа"));
            int noCloudCount = report.Count(r => r.StartsWith("[-]") && r.Contains("ОТСУТСТВУЕТ в Гугл-таблице"));
            int extraCount = report.Count(r => r.StartsWith("[+]"));

            report.Add("");
            report.Add("=== СВОДКА ===");
            report.Add($"Позиций в Заказе:                     {orderList.Count}");
            report.Add($"Позиций в ТТН:                        {ttnList.Count}");
            report.Add($"Совпало полностью:                    {okCount}");
            report.Add($"Расхождений (кол-во / цена / сумма):  {mismatchCount}");
            report.Add($"Не найдено в ТТН:                     {notFoundCount}");
            report.Add($"Нет в облачной базе:                  {noCloudCount}");
            report.Add($"Лишних позиций в ТТН:                 {extraCount}");


            int totalMissingQty = orderList
    .Where(o => !o.IsMatched)
    .Sum(o => o.Quantity);

            decimal totalMissingSum = orderList
                .Where(o => !o.IsMatched)
                .Sum(o => o.TotalSum);

            report.Add("");
            report.Add($"ИТОГО не хватает:                    {totalMissingQty} шт. на сумму {totalMissingSum:F2} руб.");

            if (mismatchCount == 0 && notFoundCount == 0 && noCloudCount == 0 && extraCount == 0)
                report.Add("\n*** ВСЁ СОШЛОСЬ! ***");
            else
                report.Add("\n*** ЕСТЬ РАСХОЖДЕНИЯ, см. выше ***");

            return report;
        }

        // =================================================================================
        // ЧТЕНИЕ ТТН
        // =================================================================================
        static List<ProductItem> ReadTtnDocument(string filePath, int startRow, int nameIdx, int qtyIdx, int priceIdx)
        {
            var result = new List<ProductItem>();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int currentRow = 0;
                    while (reader.Read())
                    {
                        currentRow++;
                        if (currentRow < startRow) continue;

                        object nameObj = reader.GetValue(nameIdx);
                        if (nameObj == null) continue;

                        string rawName = nameObj.ToString();

                        if (rawName.ToUpper().Contains("ВСЕГО") || rawName.ToUpper().Contains("ИТОГО"))
                            break;

                        int qty = ReadInt(reader.GetValue(qtyIdx));
                        decimal price = ReadDecimal(reader.GetValue(priceIdx));

                        if (qty > 0)
                        {
                            result.Add(new ProductItem
                            {
                                FullName = rawName,
                                Quantity = qty,
                                Price = price,
                                TotalSum = qty * price,
                                TtnRow = currentRow.ToString()
                            });
                        }
                    }
                }
            }
            return result;
        }

        // =================================================================================
        // ЧТЕНИЕ ЗАКАЗА
        // =================================================================================
        static List<ProductItem> ReadOrderDocument(string filePath, int startRow, int artIdx, int nameIdx, int qtyIdx, int priceIdx)
        {
            var result = new List<ProductItem>();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int currentRow = 0;
                    while (reader.Read())
                    {
                        currentRow++;
                        if (currentRow < startRow) continue;

                        object artObj = reader.GetValue(artIdx);
                        if (artObj == null) continue;

                        string art = artObj.ToString().Trim();

                        if (art.ToUpper().Contains("ИТОГО") || art.ToUpper().Contains("ВСЕГО"))
                            break;

                        string name = reader.GetValue(nameIdx)?.ToString() ?? "";

                        int qty = ReadInt(reader.GetValue(qtyIdx));
                        decimal price = ReadDecimal(reader.GetValue(priceIdx));

                        if (qty > 0)
                        {
                            result.Add(new ProductItem
                            {
                                Article = art,
                                FullName = name,
                                Quantity = qty,
                                Price = price,
                                TotalSum = qty * price,
                                OrderRow = currentRow.ToString()
                            });
                        }
                    }
                }
            }
            return result;
        }
    }
}