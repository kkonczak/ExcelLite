using ExcelLite.Internal;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;

namespace ExcelLite
{
    public static class ExcelLite
    {
        public static async Task Export<T>(string fileName, IEnumerable<T> data, CancellationToken ct = default) where T : class
        {
            using var fileStream = new FileStream(fileName, FileMode.Create);
            await Export(fileStream, data, ct);
        }

        public static Task Export<T>(Stream stream, IEnumerable<T> data, CancellationToken ct = default) where T : class =>
            Export(
                stream,
                new Workbook(
                    new Sheet[]{
                        new Sheet("Sheet 1", data)
                    }),
                ct);

        public static async Task Export(string fileName, Workbook workbook, CancellationToken ct = default)
        {
            using var fileStream = new FileStream(fileName, FileMode.Create);
            await Export(fileStream, workbook, ct);
        }

        public static async Task Export(Stream stream, Workbook workbook, CancellationToken ct = default)
        {
            using var zipArchive = new ZipArchive(stream, ZipArchiveMode.Create);
            var customFilesGenerator = new CustomFilesGenerator(workbook);

            ZipArchiveEntry contentEntry = zipArchive.CreateEntry("[Content_Types].xml");
            using (StreamWriter writer = new StreamWriter(contentEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateContentTypes());
            }

            ZipArchiveEntry relsEntry = zipArchive.CreateEntry("_rels/.rels");
            using (StreamWriter writer = new StreamWriter(relsEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateRels());
            }

            ZipArchiveEntry docPropsAppEntry = zipArchive.CreateEntry("docProps/app.xml");
            using (StreamWriter writer = new StreamWriter(docPropsAppEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateDocPropsAppXml());
            }

            ZipArchiveEntry docPropsCoreEntry = zipArchive.CreateEntry("docProps/core.xml");
            using (StreamWriter writer = new StreamWriter(docPropsCoreEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateDocPropsCoreXml());
            }

            ZipArchiveEntry workbookEntry = zipArchive.CreateEntry("xl/workbook.xml");
            using (StreamWriter writer = new StreamWriter(workbookEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateXlWorkbookXml());
            }

            ZipArchiveEntry sharedStringsEntry = zipArchive.CreateEntry("xl/sharedStrings.xml");
            using (StreamWriter writer = new StreamWriter(sharedStringsEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateXlSharedStringsXml());
            }

            ZipArchiveEntry xlRelsEntry = zipArchive.CreateEntry("xl/_rels/workbook.xml.rels");
            using (StreamWriter writer = new StreamWriter(xlRelsEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateXlRels());
            }

            ZipArchiveEntry xlThemeEntry = zipArchive.CreateEntry("xl/theme/theme1.xml");
            using (StreamWriter writer = new StreamWriter(xlThemeEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateXlTheme());
            }

            int i = 1;
            foreach (var sheet in workbook.Sheets)
            {
                ZipArchiveEntry sheetEntry = zipArchive.CreateEntry($"xl/worksheets/sheet" + i + ".xml");
                using (StreamWriter writer = new StreamWriter(sheetEntry.Open()))
                {
                    customFilesGenerator.GenerateSheet(sheet, writer, ct);
                }

                i++;
            }

            ZipArchiveEntry stylesEntry = zipArchive.CreateEntry("xl/styles.xml");
            using (StreamWriter writer = new StreamWriter(stylesEntry.Open()))
            {
                await writer.WriteAsync(customFilesGenerator.GenerateXlStylesXml());
            }
        }
    }
}