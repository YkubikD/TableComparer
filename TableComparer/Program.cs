using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ExcelDataReader;

namespace TableComparer
{
    internal class Program
    {
        // Хранилище данных для позиций из ТТН
        public class TtnProduct
        {
            public string FullName { get; set; }
            public int Quantity { get; set; }
            public decimal TotalSumWithoutNds { get; set; }
            public bool IsMatched { get; set; }
        }

        // Хранилище данных для позиций из Заказа
        public class OrderProduct
        {
            public string CustomerName { get; set; } // Наименование заказчика (столбцы 27-59)
            public string Article { get; set; }       // Артикул или код сайта (столбцы 70-77)
            public int Quantity { get; set; }
            public decimal TotalSumWithoutNds { get; set; }
            public bool IsMatched { get; set; }
        }

        [STAThread]
        static void Main(string[] args)
        {
            // Включаем поддержку кодировок для ExcelDataReader
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.WriteLine("=== ПРОГРАММА СВЕРКИ: ТТН VS ЗАКАЗ ===");

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            OpenFileDialog openFileDialog2 = new OpenFileDialog();

            openFileDialog1.Title = "Выберите файл ТТН (Накладной)";
            openFileDialog2.Title = "Выберите файл ЗАКАЗА (21vek)";
            openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls";
            openFileDialog2.Filter = "Excel Files|*.xlsx;*.xls";

            string ttnPath = ""; string orderPath = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK) ttnPath = openFileDialog1.FileName; else return;
            if (openFileDialog2.ShowDialog() == DialogResult.OK) orderPath = openFileDialog2.FileName; else return;

            Console.Clear();
            Console.WriteLine("Выполняется чтение документов, подождите...");

            // Считывание данных из файлов через подпрограммы
            var ttnList = ReadTtnDocument(ttnPath, startRow: 40);
            var orderList = ReadOrderDocument(orderPath, startRow: 25);

            Console.Clear();
            Console.WriteLine($"=== РЕЗУЛЬТАТЫ СВЕРКИ (ТТН: {ttnList.Count} поз., Заказ: {orderList.Count} поз.) ===\n");

            StringBuilder logReport = new StringBuilder();
            bool hasDiscrepancies = false;

            // Расширенный массив корней цветов Polygran (Добавлены: КРЕМ, ОПАЛ, ХЛОПОК)
            string[] colors = { "ПЕСОЧН", "СЕР", "ЧЕРН", "БЕЛ", "БЕЖ", "ГРАФИТ", "АНТРАЦИТ", "КОСМОС", "КРЕМ", "ОПАЛ", "ХЛОПОК" };

            // =================================================================================
            // // ЦИКЛ 1: Главный цикл сверки (идёт по всем товарам из файла ЗАКАЗА)
            // =================================================================================
            foreach (var ord in orderList)
            {
                TtnProduct ttnMatch = null;
                bool isDigitalCode = long.TryParse(ord.Article, out _);

                if (isDigitalCode)
                {
                    // === ПЛАН А (Для вытяжек и кодов 1С): Прямой поиск точного числового кода в названии ТТН ===
                    ttnMatch = ttnList.Find(t => !t.IsMatched && t.FullName.ToUpper().Contains(ord.Article.ToUpper()));
                }
                else
                {
                    // === ПЛАН А (Для моек): Глубокая очистка текстового артикула из Заказа ===
                    string cleanOrderArticle = ord.Article.ToUpper()
                        .Replace(" ", "").Replace("-", "").Replace("RUS", "")
                        .Replace("ЧЕРНЫЙ", "").Replace("ПЕСОЧНЫЙ", "").Replace("СЕРЫЙ", "")
                        .Replace("БЕЛЫЙ", "").Replace("БЕЖЕВЫЙ", "").Replace("КОСМОС", "")
                        .Replace("GF", "").Replace("QUARZ", "").Trim();

                    // Извлекаем маркеры сторон чаши L и R для Заказа
                    string orderSide = "БЕЗ_МАРКИРОВКИ";
                    if (cleanOrderArticle.EndsWith("L")) { orderSide = "ЛЕВАЯ"; cleanOrderArticle = cleanOrderArticle.Substring(0, cleanOrderArticle.Length - 1); }
                    else if (cleanOrderArticle.EndsWith("R")) { orderSide = "ПРАВАЯ"; cleanOrderArticle = cleanOrderArticle.Substring(0, cleanOrderArticle.Length - 1); }

                    cleanOrderArticle = cleanOrderArticle.Replace("M", "М");

                    if (cleanOrderArticle.Length >= 2)
                    {
                        ttnMatch = ttnList.Find(t => {
                            if (t.IsMatched) return false;
                            string cleanTtnName = t.FullName.ToUpper().Replace(" ", "").Replace("-", "").Replace("M", "М").Replace("GF", "").Replace("QUARZ", "");
                            if (!cleanTtnName.Contains(cleanOrderArticle)) return false;

                            // Контроль стороны чаши в ТТН
                            string ttnSide = "БЕЗ_МАРКИРОВКИ";
                            if (cleanTtnName.Contains("ЛЕВ") || cleanTtnName.Contains("СЛЕВА")) ttnSide = "ЛЕВАЯ";
                            if (cleanTtnName.Contains("ПРАВ") || cleanTtnName.Contains("СПРАВА")) ttnSide = "ПРАВАЯ";

                            return orderSide == ttnSide;
                        });
                    }
                }

                // === ПЛАН Б (РЕЗЕРВНЫЙ ПО ЦВЕТУ И МОДЕЛИ): Если План А не сработал (для Polygran и Granfest) ===
                if (ttnMatch == null)
                {
                    string nameUpper = ord.CustomerName.ToUpper();
                    string extractedModel = "";

                    // Добавили новые модели моек со скриншота: ARGO 460, 560, 760
                    if (nameUpper.Contains("BLADE 500")) extractedModel = "BLADE500";
                    else if (nameUpper.Contains("TOLERO 580")) extractedModel = "TOLERO580";
                    else if (nameUpper.Contains("ARGO 460")) extractedModel = "ARGO460";
                    else if (nameUpper.Contains("ARGO 560")) extractedModel = "ARGO560";
                    else if (nameUpper.Contains("ARGO 760")) extractedModel = "ARGO760";
                    else if (nameUpper.Contains("GALS 862")) extractedModel = "GALS862";
                    else if (nameUpper.Contains("BRIG 772") || nameUpper.Contains("BRIG-772")) extractedModel = "BRIG772";
                    else if (nameUpper.Contains("QUARZ 18") || nameUpper.Contains("Z18") || nameUpper.Contains("Z-18")) extractedModel = "QUARZ18";
                    else if (nameUpper.Contains("РОТОНДА")) extractedModel = "РОТОНДА";
                    else if (nameUpper.Contains("ПЕРГОЛА")) extractedModel = "ПЕРГОЛА";
                    else if (nameUpper.Contains("SLIDE")) extractedModel = "SLIDE";
                    else
                    {
                        int polyIdx = nameUpper.IndexOf("POLYGRAN");
                        if (polyIdx != -1 && polyIdx + 9 < nameUpper.Length)
                        {
                            string sub = nameUpper.Substring(polyIdx + 9).Trim();
                            string[] words = sub.Split(' ');
                            if (words.Length >= 2) extractedModel = words[0] + words[1];
                        }
                    }
                    extractedModel = extractedModel.Replace(" ", "").Replace("-", "");

                    string detectedColor = "";
                    foreach (var c in colors) { if (nameUpper.Contains(c)) { detectedColor = c; break; } }

                    if (!string.IsNullOrEmpty(extractedModel))
                    {
                        ttnMatch = ttnList.Find(t => {
                            if (t.IsMatched) return false;

                            // Очитка названия ТТН от знаков № и лишних символов
                            string cleanTtn = t.FullName.ToUpper().Replace(" ", "").Replace("-", "").Replace("№", "");

                            bool modelOk = cleanTtn.Contains(extractedModel) || (extractedModel == "QUARZ18" && cleanTtn.Contains("Z18"));
                            bool colorOk = string.IsNullOrEmpty(detectedColor) || cleanTtn.Contains(detectedColor);
                            return modelOk && colorOk;
                        });
                    }
                }

                // === СВЕРКА МАТЕМАТИКИ И ВЫВОД РАСХОЖДЕНИЙ ===
                if (ttnMatch != null)
                {
                    ord.IsMatched = true; ttnMatch.IsMatched = true;
                    bool qtyMismatch = ord.Quantity != ttnMatch.Quantity;
                    bool sumMismatch = Math.Round(ord.TotalSumWithoutNds, 2) != Math.Round(ttnMatch.TotalSumWithoutNds, 2);

                    if (qtyMismatch || sumMismatch)
                    {
                        hasDiscrepancies = true;
                        string err = $"[!] Расхождение по товару (Идентификатор: {ord.Article})\n" +
                                     $"    Имя в Заказе: {ord.CustomerName}\n" +
                                     $"    Имя в ТТН:    {ttnMatch.FullName}\n";
                        if (qtyMismatch) err += $"    [-] КОЛИЧЕСТВО: Заказ = {ord.Quantity} шт. | ТТН = {ttnMatch.Quantity} шт. (Разница: {ord.Quantity - ttnMatch.Quantity})\n";
                        if (sumMismatch) err += $"    [-] СУММА БЕЗ НДС: Заказ = {ord.TotalSumWithoutNds:F2} руб. | ТТН = {ttnMatch.TotalSumWithoutNds:F2} руб.\n";

                        Console.ForegroundColor = ConsoleColor.Yellow; Console.WriteLine(err.TrimEnd()); Console.ResetColor();
                        Console.WriteLine(new string('-', 70)); logReport.AppendLine(err + new string('-', 50) + "\n");
                    }
                }
                else
                {
                    hasDiscrepancies = true;
                    string err = $"[-] Товар из Заказа [{ord.Article}] ({ord.CustomerName}) ОТСУТСТВУЕТ в ТТН!\n";
                    Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine(err.TrimEnd()); Console.ResetColor();
                    Console.WriteLine(new string('-', 70)); logReport.AppendLine(err + new string('-', 50) + "\n");
                }
            }
            // =================================================================================
            // КОНЕЦ // ЦИКЛ 1
            // =================================================================================

            // =================================================================================
            // // ЦИКЛ 2: Проверка лишних позиций в ТТН (ищет товары, к которым не подошел заказ)
            // =================================================================================
            foreach (var ttn in ttnList)
            {
                if (!ttn.IsMatched)
                {
                    hasDiscrepancies = true;
                    string err = $"[+] Позиция из ТТН [{ttn.FullName}] отсутствует в Заказе (лишний товар или не распознан)!\n";
                    Console.ForegroundColor = ConsoleColor.Cyan; Console.WriteLine(err.TrimEnd()); Console.ResetColor();
                    Console.WriteLine(new string('-', 70)); logReport.AppendLine(err + new string('-', 50) + "\n");
                }
            }
            // =================================================================================
            // КОНЕЦ // ЦИКЛ 2
            // =================================================================================

            if (!hasDiscrepancies)
            {
                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("✔ Сверка успешно завершена! Расхождений не обнаружено."); Console.ResetColor();
            }
            else
            {
                try
                {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    File.WriteAllText(Path.Combine(desktopPath, "Результат_Сверки.txt"), logReport.ToString(), Encoding.UTF8);
                    Console.ForegroundColor = ConsoleColor.Blue; Console.WriteLine("\n[!] Подробный лог сохранен на Рабочий стол: 'Результат_Сверки.txt'"); Console.ResetColor();
                }
                catch { }
            }

            Console.WriteLine("\nНажмите Enter для завершения работы...");
            Console.ReadLine();
        } // Конец метода Main

        /// <summary>
        /// Чтение ТТН: старт с 40 строки, лимит по «ИТОГО»
        /// </summary>
        static List<TtnProduct> ReadTtnDocument(string filePath, int startRow)
        {
            var list = new List<TtnProduct>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int currentRow = 0;

                // =============================================================================
                // // ЦИКЛ 3: Построчное чтение и разбор файла НАКЛАДНОЙ (ТТН)
                // =============================================================================
                while (reader.Read())
                {
                    currentRow++;
                    if (currentRow < startRow) continue;
                    if (reader.FieldCount < 25) continue; // Защита от коротких сервисных строк

                    // Наименование товара (Столбцы 1-14, индексы в C# 0-13)
                    string rawName = "";
                    for (int i = 0; i <= 13; i++)
                    {
                        var val = reader.GetValue(i);
                        if (val != null && !string.IsNullOrEmpty(val.ToString()))
                        {
                            rawName = val.ToString().Trim();
                            break;
                        }
                    }

                    // Жесткое ограничение ТТН: Стоп на итоговых строках
                    if (string.IsNullOrEmpty(rawName) ||
                        rawName.StartsWith("ИТОГО", StringComparison.OrdinalIgnoreCase) ||
                        rawName.StartsWith("Всего", StringComparison.OrdinalIgnoreCase) ||
                        rawName.Contains("Отпуск разрешил") ||
                        rawName.Contains("Сдал грузоотправитель"))
                    {
                        if (rawName.StartsWith("ИТОГО", StringComparison.OrdinalIgnoreCase) || rawName.StartsWith("Всего", StringComparison.OrdinalIgnoreCase))
                            break; // Полностью выходим из цикла чтения ТТН
                        continue;
                    }

                    // Количество (Столбцы 18-20, индексы в C# 17-19)
                    int qty = 0;
                    for (int i = 17; i <= 19; i++)
                    {
                        var val = reader.GetValue(i);
                        if (val != null)
                        {
                            int.TryParse(val.ToString(), out qty);
                            if (qty > 0) break;
                        }
                    }

                    // Цена за 1 шт (Столбцы 21-24, индексы в C# 20-23)
                    decimal price = 0;
                    for (int i = 20; i <= 23; i++)
                    {
                        var val = reader.GetValue(i);
                        if (val != null)
                        {
                            decimal.TryParse(val.ToString(), out price);
                            if (price > 0) break;
                        }
                    }

                    if (qty > 0)
                    {
                        list.Add(new TtnProduct
                        {
                            FullName = rawName.Replace("\n", " ").Replace("\r", ""),
                            Quantity = qty,
                            TotalSumWithoutNds = qty * price // Вычисляем стоимость всей строки без НДС
                        });
                    }
                }
                // =============================================================================
                // КОНЕЦ // ЦИКЛ 3
                // =============================================================================
            }
            return list;
        }

        /// <summary>
        /// Чтение Заказа: старт с 25 строки, лимит по пустой строке артикула
        /// </summary>
        static List<OrderProduct> ReadOrderDocument(string filePath, int startRow)
        {
            var list = new List<OrderProduct>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int currentRow = 0;

                // =============================================================================
                // // ЦИКЛ 4: Построчное чтение и разбор файла ЗАКАЗА ПОСТАВЩИКУ
                // =============================================================================
                while (reader.Read())
                {
                    currentRow++;
                    if (currentRow < startRow) continue;
                    if (reader.FieldCount < 125) continue; // Защита по ширине таблицы Заказа

                    // 1. Артикул: Столбцы 70-77 (Индекс в C# строго 69)
                    object artObj = reader.GetValue(69);

                    // ЖЕСТКОЕ ОГРАНИЧЕНИЕ ЗАКАЗА: Если ячейка артикула пуста - прекращаем чтение файла
                    if (artObj == null || string.IsNullOrEmpty(artObj.ToString().Trim()))
                    {
                        break; // Выходим из цикла чтения Заказа, так как позиции закончились
                    }

                    string article = artObj.ToString().Trim();
                    if (article == "Артикул") continue; // Пропуск повторного заголовка

                    // 2. Наименование заказчика (Столбцы 27-59, индекс в C# строго 26)
                    object custNameObj = reader.GetValue(26);
                    string customerName = custNameObj != null ? custNameObj.ToString().Trim() : "";

                    // 3. Количество: Левая половинка блока (Столбцы 102-105, индекс в C# строго 101)
                    object qtyObj = reader.GetValue(101);
                    int qty = 0;
                    if (qtyObj != null) int.TryParse(qtyObj.ToString(), out qty);

                    // 4. Цена закупки без НДС: Столбцы 120-124 (индекс в C# строго 119)
                    object priceObj = reader.GetValue(119);
                    decimal price = 0;
                    if (priceObj != null) decimal.TryParse(priceObj.ToString(), out price);

                    if (qty > 0)
                    {
                        list.Add(new OrderProduct
                        {
                            Article = article,
                            CustomerName = customerName,
                            Quantity = qty,
                            TotalSumWithoutNds = qty * price // Вычисляем стоимость всей строки без НДС
                        });
                    }
                }
                // =============================================================================
                // КОНЕЦ // ЦИКЛ 4
                // =============================================================================
            }
            return list;
        }
    } // Конец класса Program
} // Конец namespace TableComparer