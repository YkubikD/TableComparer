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

            Console.WriteLine("[3/3] Выполнение автоматической сверки...");
            List<string> reportLines = ExecuteComparison(ttnList, orderList);

            Console.WriteLine("\n=== РЕЗУЛЬТАТЫ СВЕРКИ ===");
            foreach (var line in reportLines)
            {
                if (line.StartsWith("[OK]")) Console.ForegroundColor = ConsoleColor.Green;
                else if (line.StartsWith("[-]")) Console.ForegroundColor = ConsoleColor.Red;
                else if (line.StartsWith("[!]")) Console.ForegroundColor = ConsoleColor.Yellow;
                else if (line.StartsWith("[+]")) Console.ForegroundColor = ConsoleColor.Magenta;

                Console.WriteLine(line);
                Console.ResetColor();
            }
            Console.WriteLine("=========================\n");

            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("Готово. Нажмите Enter для выхода...");
            Console.ReadLine();
        }

        static bool LoadGoogleSpreadsheet(string url)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Скачиваем CSV поток данных
                    var stream = client.GetStreamAsync(url).Result;

                    // Используем встроенный неубиваемый TextFieldParser для разбора CSV с кавычками
                    using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(stream, Encoding.UTF8))
                    {
                        parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                        parser.SetDelimiters(",");
                        parser.HasFieldsEnclosedInQuotes = true; // Защита от запятых внутри названий моек

                        bool isHeader = true;

                        while (!parser.EndOfData)
                        {
                            string[] parts = parser.ReadFields();
                            if (isHeader) { isHeader = false; continue; } // Пропускаем шапку

                            if (parts == null || parts.Length < 2) continue;

                            string fullName = parts[0]?.Trim() ?? ""; // Колонка A
                            string article = parts[1]?.Trim() ?? "";  // Колонка B

                            // СТОП-МАРКЕР: Если артикул пустой — таблица гарантированно закончилась
                            if (string.IsNullOrEmpty(article) || article.ToLower().Contains("штамп"))
                                break;

                            // Если артикул есть, но имя пустое — подставляем артикул, чтобы не упасть
                            if (string.IsNullOrEmpty(fullName))
                            {
                                fullName = "Модель без имени (" + article + ")";
                            }

                            // Пропускаем служебные заголовки
                            if (article.ToLower().Contains("артикул"))
                                continue;

                            // Создаем чистый ключ без пробелов и дефисов
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




        //конец части 1




        // Основной метод сверки
        static List<string> ExecuteComparison(List<ProductItem> ttnList, List<ProductItem> orderList)
        {
            List<string> report = new List<string>();
            bool hasDiscrepancies = false;

            // Проходим по каждой позиции Заказа от "21 века"
            foreach (var orderItem in orderList)
            {
                // Очищаем артикул из заказа для поиска в облачном словаре
                string cleanOrderArt = orderItem.Article.Replace(" ", "").Replace("-", "").ToUpper();

                // Проверяем, внесли ли вы этот артикул в Гугл-таблицу (Лист 4)
                if (!GoogleCloudDictionary.ContainsKey(cleanOrderArt))
                {
                    hasDiscrepancies = true;
                    report.Add($"[!] ПРЕДУПРЕЖДЕНИЕ: Артикул [{orderItem.Article}] ({orderItem.FullName}) ОТСУТСТВУЕТ в Гугл-таблице! Внесите его в базу.");
                    continue;
                }

                // Достаем из облачной базы полное текстовое наименование
                string cloudTtnNameTarget = GoogleCloudDictionary[cleanOrderArt].ToUpper().Replace(" ", "").Replace("-", "");

                // Выделяем базовое имя модели для сопоставления (например, ARGO420, BLADE500, S416)
                string modelKey = "";
                if (cloudTtnNameTarget.Contains("ARGO")) modelKey = "ARGO";
                else if (cloudTtnNameTarget.Contains("BLADE")) modelKey = "BLADE";
                else if (cloudTtnNameTarget.Contains("TOLERO")) modelKey = "TOLERO";
                else if (cloudTtnNameTarget.Contains("BRIG")) modelKey = "BRIG";
                else if (cloudTtnNameTarget.Contains("GALS")) modelKey = "GALS";
                else if (cloudTtnNameTarget.Contains("SMART")) modelKey = "SMART";
                else if (cloudTtnNameTarget.Contains("URBAN")) modelKey = "URBAN";
                else if (cloudTtnNameTarget.Contains("CORNER")) modelKey = "CORNER";
                else if (cloudTtnNameTarget.Contains("RONDO")) modelKey = "RONDO";
                else if (cloudTtnNameTarget.Contains("UNIQUE")) modelKey = "UNIQUE";
                else if (cloudTtnNameTarget.Contains("QUADRO")) modelKey = "QUADRO";

                // Если в облачном наименовании есть специфичный цифровой индекс модели (например, 420, 460, 445)
                string digits = new string(cloudTtnNameTarget.Where(char.IsDigit).ToArray());

                // Ищем подходящий товар в ТТН отгрузки склада среди еще не сопоставленных
                var matchedTtn = ttnList.Find(t => !t.IsMatched && (
                    // Вариант 1: Полное совпадение очищенного текста из облака с текстом ТТН
                    t.FullName.ToUpper().Replace(" ", "").Replace("-", "").Contains(cloudTtnNameTarget) ||
                    // Вариант 2: Защитный поиск по бренду, цифрам серии и стороне чаши
                    (!string.IsNullOrEmpty(modelKey) &&
                     t.FullName.ToUpper().Replace(" ", "").Replace("-", "").Contains(modelKey) &&
                     (!string.IsNullOrEmpty(digits) && t.FullName.Contains(digits)) &&
                     MatchSides(t.FullName, orderItem.FullName))
                ));

                if (matchedTtn != null)
                {
                    matchedTtn.IsMatched = true;
                    orderItem.IsMatched = true;

                    // Проверяем расхождения по количеству или по цене отгрузки
                    if (matchedTtn.Quantity != orderItem.Quantity || Math.Abs(matchedTtn.Price - orderItem.Price) > 0.05m)
                    {
                        hasDiscrepancies = true;
                        string diff = $"[!] Расхождение по товару (Код 1С: {orderItem.Article}):\n" +
                                      $"    Заказ (21 век): {orderItem.FullName}\n" +
                                      $"    ТТН (Отгрузка): {matchedTtn.FullName}\n";

                        if (matchedTtn.Quantity != orderItem.Quantity)
                            diff += $"    [-] КОЛИЧЕСТВО: Заказ = {orderItem.Quantity} шт. | ТТН = {matchedTtn.Quantity} шт. (Разница: {orderItem.Quantity - matchedTtn.Quantity})\n";

                        if (Math.Abs(matchedTtn.Price - orderItem.Price) > 0.05m)
                            diff += $"    [-] ЦЕНА: Заказ = {orderItem.Price} руб. | ТТН = {matchedTtn.Price} руб.\n";

                        report.Add(diff);
                    }
                }
                else
                {
                    hasDiscrepancies = true;
                    report.Add($"[-] Товар из Заказа [{orderItem.Article}] ({orderItem.FullName}) НЕ НАЙДЕН в ТТН отгрузки!");
                }
            }

            // Ищем позиции, которые склад отгрузил по ТТН, но "21 век" их не заказывал
            foreach (var ttnItem in ttnList.Where(t => !t.IsMatched))
            {
                hasDiscrepancies = true;
                report.Add($"[+] Лишний товар в ТТН! Позиция [{ttnItem.FullName}] отсутствует в Заказе клиента.");
            }

            if (!hasDiscrepancies) report.Add("[OK] Склад полностью и правильно отгрузил Заказ. Позиции, количества и цены совпали идеально.");
            return report;
        }

        // Вспомогательный метод проверки направления чаши (ЛЕВ/ПРАВ/L/R)
        static bool MatchSides(string ttnName, string orderName)
        {
            string t = ttnName.ToUpper();
            string o = orderName.ToUpper();

            bool ttnSide = t.Contains("ЛЕВ") || t.Contains("L") || t.Contains("СЛЕВА") || t.Contains("ЛЕВАЯ");
            bool ordSide = o.Contains("ЛЕВ") || o.Contains("L") || o.Contains("СЛЕВА") || o.Contains("ЛЕВАЯ");

            return ttnSide == ordSide;
        }





        //конец части 2




        // Метод для чтения Накладной (ТТН)
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

                        // Прерываем чтение при обнаружении итоговых строк
                        if (rawName.ToUpper().Contains("ВСЕГО") || rawName.ToUpper().Contains("ИТОГО"))
                            break;

                        int qty = 0;
                        object qtyObj = reader.GetValue(qtyIdx);
                        if (qtyObj != null) int.TryParse(qtyObj.ToString(), out qty);

                        decimal price = 0;
                        object priceObj = reader.GetValue(priceIdx);
                        if (priceObj != null) decimal.TryParse(priceObj.ToString(), out price);

                        if (qty > 0)
                        {
                            result.Add(new ProductItem
                            {
                                FullName = rawName,
                                Quantity = qty,
                                Price = price,
                                TtnRow = currentRow.ToString()
                            });
                        }
                    }
                }
            }
            return result;
        }

        // Метод для чтения файла Заказа
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

                        // Маркер окончания данных в заказе
                        if (art.ToUpper().Contains("ИТОГО") || art.ToUpper().Contains("ВСЕГО"))
                            break;

                        string name = reader.GetValue(nameIdx)?.ToString() ?? "";

                        int qty = 0;
                        object qtyObj = reader.GetValue(qtyIdx);
                        if (qtyObj != null) int.TryParse(qtyObj.ToString(), out qty);

                        decimal price = 0;
                        object priceObj = reader.GetValue(priceIdx);
                        if (priceObj != null) decimal.TryParse(priceObj.ToString(), out price);

                        if (qty > 0)
                        {
                            result.Add(new ProductItem
                            {
                                Article = art,
                                FullName = name,
                                Quantity = qty,
                                Price = price,
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

//конец часть 3