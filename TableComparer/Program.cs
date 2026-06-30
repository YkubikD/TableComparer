using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDataReader;

namespace TableComparer
{
    internal class Program
    {
        class ProductInfo
        {
            public int Quantity { get; set; }
            public decimal Price { get; set; }
        }

        [STAThread]
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Console.Title = "Автоматическая тройная сверка (Модель + Кол-во + Цена)";
            Console.WriteLine("=== ПРОГРАММА ТРОЙНОЙ СВЕРКИ НАКЛАДНЫХ И ЗАКАЗОВ ===");
            Console.WriteLine();

            Console.WriteLine("Выберите первый файл (Накладную из 1С)...");
            string pathFile1 = OpenFileDialog("Выберите файл Накладной");
            if (string.IsNullOrEmpty(pathFile1)) return;

            Console.WriteLine("Выберите второй файл (Заказ поставщику)...");
            string pathFile2 = OpenFileDialog("Выберите файл Заказа");
            if (string.IsNullOrEmpty(pathFile2)) return;

            Console.WriteLine("Загрузка и перекрестный анализ данных...");

            try
            {
                // Читаем Накладную: старт с 40 строки, имя=0, колво=17, цена=20
                var dataFile1 = ReadExcelDocument(pathFile1, startRow: 40, nameIdx: 0, qtyIdx: 17, priceIdx: 20);

                // Читаем Заказ: старт с 25 строки, артикул=69, колво=101, цена=119
                var dataFile2 = ReadExcelDocument(pathFile2, startRow: 25, nameIdx: 69, qtyIdx: 101, priceIdx: 119);

                Console.WriteLine($"\nУспешно считано из Накладной: {dataFile1.Count} позиций.");
                Console.WriteLine($"Успешно считано из Заказа:    {dataFile2.Count} позиций.");
                Console.WriteLine("\n--- РЕЗУЛЬТАТ ТРОЙНОЙ СВЕРКИ (ОТКЛОНЕНИЯ) ---");

                bool hasDiscrepancies = false;

                // Собираем все уникальные артикулы из обоих файлов
                var allKeys = new HashSet<string>(dataFile1.Keys);
                allKeys.UnionWith(dataFile2.Keys);

                foreach (var art in allKeys)
                {
                    var p1 = dataFile1.ContainsKey(art) ? dataFile1[art] : new ProductInfo { Quantity = 0, Price = 0 };
                    var p2 = dataFile2.ContainsKey(art) ? dataFile2[art] : new ProductInfo { Quantity = 0, Price = 0 };

                    // Проверяем расхождение по количеству или цене за штуку без НДС
                    if (p1.Quantity != p2.Quantity || p1.Price != p2.Price)
                    {
                        hasDiscrepancies = true;
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"❌ Расхождение по модели: {art}");
                        Console.ResetColor();

                        if (p2.Quantity == 0)
                        {
                            Console.WriteLine($"   Лишний товар в Накладной! Нет в Заказе. По факту: {p1.Quantity} шт. по {p1.Price} руб.");
                        }
                        else if (p1.Quantity == 0)
                        {
                            Console.WriteLine($"   Недовоз! Есть в Заказе ({p2.Quantity} шт. по {p2.Price} руб.), но отсутствует в Накладной.");
                        }
                        else
                        {
                            if (p1.Quantity != p2.Quantity)
                            {
                                Console.WriteLine($"   ⚠ КОЛИЧЕСТВО не сошлось: В Заказе {p2.Quantity} шт. | В Накладной {p1.Quantity} шт. (Разница: {p1.Quantity - p2.Quantity} шт.)");
                            }
                            if (p1.Price != p2.Price)
                            {
                                Console.WriteLine($"   ⚠ ЦЕНА БЕЗ НДС не сошлась: В Заказе {p2.Price} руб. | В Накладной {p1.Price} руб. (Разница: {p1.Price - p2.Price} руб.)");
                            }
                        }
                        Console.WriteLine();
                    }
                }

                if (!hasDiscrepancies && allKeys.Count > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("✅ Идеально! Все наименования, количества и цены без НДС сошлись на 100%.");
                    Console.ResetColor();
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Ошибка при обработке файлов: {ex.Message}");
                Console.ResetColor();
            }

            Console.WriteLine("\nНажмите любую клавишу для выхода...");
            Console.ReadKey();
        }

        static string OpenFileDialog(string title)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = title;
                ofd.Filter = "Excel|*.xlsx;*.xls";
                if (ofd.ShowDialog() == DialogResult.OK) return ofd.FileName;
            }
            return null;
        }

        static Dictionary<string, ProductInfo> ReadExcelDocument(string filePath, int startRow, int nameIdx, int qtyIdx, int priceIdx)
        {
            var result = new Dictionary<string, ProductInfo>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int currentRow = 0;
                while (reader.Read())
                {
                    currentRow++;
                    if (currentRow < startRow) continue;

                    if (reader.FieldCount <= Math.Max(nameIdx, Math.Max(qtyIdx, priceIdx))) continue;

                    object nameObj = reader.GetValue(nameIdx);
                    if (nameObj == null) continue;

                    string rawName = nameObj.ToString().Trim();

                    // Фильтр технических подписей 1С
                    if (string.IsNullOrEmpty(rawName) || rawName.StartsWith("ИТОГО") || rawName.StartsWith("Всего") ||
                        rawName.Contains("Отпуск разрешил") || rawName.Contains("Сдал грузоотправитель") ||
                        rawName.Contains("Всего наименований")) continue;

                    string cleanKey = ExtractCleanKey(rawName);
                    if (string.IsNullOrEmpty(cleanKey)) continue;

                    // Учитываем сторону чаши для предотвращения пересортицы
                    if (rawName.ToLower().Contains("лев")) cleanKey += " (левая)";
                    if (rawName.ToLower().Contains("прав")) cleanKey += " (правая)";

                    // Читаем количество штук
                    object qtyObj = reader.GetValue(qtyIdx);
                    int qty = 0;
                    if (qtyObj != null) int.TryParse(qtyObj.ToString(), out qty);

                    // Читаем цену за штуку без НДС
                    object priceObj = reader.GetValue(priceIdx);
                    decimal price = 0;
                    if (priceObj != null) decimal.TryParse(priceObj.ToString(), out price);

                    if (qty > 0)
                    {
                        if (result.ContainsKey(cleanKey))
                            result[cleanKey].Quantity += qty;
                        else
                            result[cleanKey] = new ProductInfo { Quantity = qty, Price = price };
                    }
                }
            }
            return result;
        }

        static string ExtractCleanKey(string text)
        {
            text = text.ToUpper().Trim();

            // Если это вытяжка или фильтр
            if (text.Contains("SLIDE")) return "ВЫТЯЖКА-ELIKOR";
            if (text.Contains("АКВАОСМОС")) return "ФИЛЬТР-АКВАОСМОС";

            // Мягко ищем ключевую модель в тексте (заменяем английские M/H на русские и убираем пробелы)
            string[] models = { "S-404", "S 404", "S-417", "S 417", "S-409", "S 409", "S-420", "S 420", "S-416", "S 416",
                                "EC-245", "EC 245", "EC-445M", "EC-445М", "EC 445 M", "EC 445 М", "EC-445", "EC 445",
                                "EC-220M", "EC-220М", "EC 220 M", "EC 220 М", "EC-257", "EC 257", "EC-296", "EC 296", "N-205", "N 205",
                                "GR-460", "GR 460", "GR-555", "GR 555", "GR-770", "GR 770", "GALS-760", "GALS 760", "ARGO-560", "ARGO 560" };

            foreach (var model in models)
            {
                if (text.Contains(model))
                {
                    // Приводим все к единому стандарту: без пробелов, с дефисом и русской буквой М
                    return model.Replace(" ", "-").Replace("H", "M");
                }
            }

            return text.Length > 15 ? text.Substring(0, 15) : text;
        }
    }
}