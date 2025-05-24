using System.IO.Compression;
using System.Xml;
using SPSReader.utils;
using SPSReader.types;
namespace SPSReader
{
    
    public static class XlReader
    {
        public static void Main()
        {
            var c = ExcelReaderAsync("/home/manu/Escritorio/xmls/Excels/sueldo.xlsx").Result;
            var sheets = c.Sheets[0];
            var sheet = sheets["sueldo"];  // Ahora se usa el nombre "sueldo" en lugar de "sheet1.xml"
            foreach (var sht in sheet)
            {
                foreach (var s in sht.Keys)
                {
                    Console.WriteLine($"{s} - {sht[s]}");
                }
            }
        }

        public static async Task<XlSheets> ExcelReaderAsync(string filePath)
        {
            var spreadSheets = new XlSheets();
            var xlSheets = spreadSheets.Sheets;
            var settings = new XmlReaderSettings { Async = true };
            var doc = new XmlDocument();

            // Abrimos el archivo Zip
            using var zipArchive = ZipFile.Open(filePath, ZipArchiveMode.Read);

            // Procesar el workbook rels para obtener las relaciones
            var workbookRelsEntry = zipArchive.GetEntry("xl/_rels/workbook.xml.rels");
            var workbookRels = new Dictionary<string, string>();
            if (workbookRelsEntry != null)
            {
                workbookRels = await ProcessWorkbookRelsAsync(workbookRelsEntry, settings);
            }

            // Procesar el workbook para obtener los nombres de las hojas
            var workbookEntry = zipArchive.GetEntry("xl/workbook.xml");
            var sheetNames = new Dictionary<string, string>();
            if (workbookEntry != null)
            {
                sheetNames = await ProcessWorkbookAsync(workbookEntry, workbookRels, settings);
            }

            // Procesar las hojas de cálculo en xl/worksheets/
            foreach (var zipArchiveEntry in zipArchive.Entries)
            {
                if (zipArchiveEntry.FullName.StartsWith("xl/worksheets/") && zipArchiveEntry.FullName.EndsWith(".xml"))
                {
                    var sheetFileName = zipArchiveEntry.FullName.Split('/').Last();  // Obtenemos el nombre del archivo (ejemplo: sheet1.xml)

                    // Buscar el nombre de la hoja en workbook.xml basado en su r:id
                    var sheetName = sheetNames.FirstOrDefault(kv => kv.Value == sheetFileName).Key;

                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        // Procesar el archivo sheet XML correspondiente
                        await ProcessSheetXmlAsync(zipArchiveEntry, spreadSheets, settings, doc, sheetName);
                    }
                }
            }

            return spreadSheets;
        }

        private static async Task<Dictionary<string, string>> ProcessWorkbookRelsAsync(ZipArchiveEntry workbookRelsEntry, XmlReaderSettings settings)
        {
            var workbookRels = new Dictionary<string, string>();

            await using var stream = workbookRelsEntry.Open();
            using var reader = XmlReader.Create(stream, settings);

            while (await reader.ReadAsync())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "Relationship")
                {
                    var id = reader.GetAttribute("Id");
                    var target = reader.GetAttribute("Target");
                    if (target != null && target.Contains("worksheets"))
                    {
                        // Asociamos el Id con el archivo de la hoja en target (ejemplo: worksheets/sheet1.xml)
                        workbookRels[id] = target.Split('/').Last(); // Obtenemos el nombre del archivo (sheet1.xml)
                    }
                }
            }

            return workbookRels;
        }

        private static async Task<Dictionary<string, string>> ProcessWorkbookAsync(ZipArchiveEntry workbookEntry, Dictionary<string, string> workbookRels, XmlReaderSettings settings)
        {
            var sheetNames = new Dictionary<string, string>();

            await using var stream = workbookEntry.Open();
            using var reader = XmlReader.Create(stream, settings);

            while (await reader.ReadAsync())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "sheet")
                {
                    var name = reader.GetAttribute("name");
                    var rId = reader.GetAttribute("r:id");
                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(rId))
                    {
                        // Relacionamos el nombre de la hoja con su archivo XML usando r:id
                        if (workbookRels.TryGetValue(rId, out var targetSheet))
                        {
                            sheetNames[name] = targetSheet;  // Ejemplo: {"sueldo" -> "sheet1.xml"}
                        }
                    }
                }
            }

            return sheetNames;
        }

        private static async Task ProcessSheetXmlAsync(ZipArchiveEntry sheetEntry, XlSheets spreadSheets, XmlReaderSettings settings, XmlDocument doc, string sheetName)
        {
            await using var stream = sheetEntry.Open();
            using var reader = XmlReader.Create(stream, settings);

            var cellValues = new List<Dictionary<string, string>>();

            while (await XlHelpers.ReadToFollowingAsync(reader, "sheetData"))
            {
                if (!await XlHelpers.ReadToDescendantAsync(reader, "row")) continue;

                do
                {
                    var xmlNode = doc.ReadNode(reader);
                    var nodeChildren = xmlNode?.ChildNodes;
                    if (GetInnerText(nodeChildren, out var rowCells))
                    {
                        cellValues.Add(rowCells);
                    }
                } while (reader is { NodeType: XmlNodeType.Element, Name: "row" });

                // Guardamos los valores bajo el nombre de la hoja (ejemplo: "sueldo")
                var xlSheet = spreadSheets.Sheet;
                xlSheet[sheetName ?? throw new InvalidOperationException()] = cellValues;
                spreadSheets.Sheets.Add(xlSheet);
            }
        }

        // Función utilitaria para extraer los valores de las celdas
        public static bool GetInnerText(XmlNodeList? nodes, out Dictionary<string, string> rowCells)
        {
            rowCells = new Dictionary<string, string>();
            if (nodes == null) return false;

            foreach (XmlNode node in nodes)
            {
                var innerText = node.InnerText;
                if (string.IsNullOrEmpty(innerText)) continue;
                var cellKey = $"Cell_{rowCells.Count + 1}";
                rowCells[cellKey] = innerText;
            }
            return rowCells.Count > 0;
        }
    }
}
