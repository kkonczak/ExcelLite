using ExcelLite.Attributes;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace ExcelLite.Internal
{
    public class CustomFilesGenerator
    {
        private Workbook _workbook;
        private List<ExcelCellFormat> _excelCellFormatList = new List<ExcelCellFormat>();

        public CustomFilesGenerator(Workbook workbook)
        {
            _workbook = workbook;
        }

        public string GenerateDocPropsAppXml()
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>ExcelLite</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Sheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{{_workbook.Sheets.Count()}}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="{{_workbook.Sheets.Count()}}" baseType="lpstr">
            """);

            foreach (var sheet in _workbook.Sheets)
            {
                if (sheet.Name.Length >= 32)
                {
                    throw new ArgumentException("Sheet name is too long!");
                }

                stringBuilder.Append("<vt:lpstr>");
                stringBuilder.Append(sheet.Name);
                stringBuilder.Append("</vt:lpstr>");
            }

            stringBuilder.Append($$"""
            </vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>15.0300</AppVersion></Properties>
            """);

            return stringBuilder.ToString();
        }

        public string GenerateDocPropsCoreXml() =>
            $$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>
            """;

        public string GenerateContentTypes()
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.Append($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            """);

            var i = 1;
            foreach (var sheet in _workbook.Sheets)
            {
                stringBuilder.Append("<Override PartName=\"/xl/worksheets/sheet");
                stringBuilder.Append((i++).ToString());

                stringBuilder.Append($$"""
                .xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                """);
            }

            stringBuilder.Append($$"""
            <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>
            """);

            return stringBuilder.ToString();
        }

        public string GenerateRels() =>
            $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
            """;

        public string GenerateXlWorkbookXml()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/><workbookPr defaultThemeVersion="153222"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="C:\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="12084" windowHeight="4596"/></bookViews><sheets>
            """);

            var i = 1;
            foreach (var sheet in _workbook.Sheets)
            {
                var visibility = sheet.Visibility switch
                {
                    SheetVisibility.Hidden => " state=\"hidden\"",
                    SheetVisibility.VeryHidden => " state=\"veryHidden\"",
                    _ => string.Empty
                };
                stringBuilder.Append($"<sheet name=\"{sheet.Name}\" sheetId=\"{i}\" r:id=\"rId{i}\"{visibility}/>");
                i++;
            }

            stringBuilder.Append($$"""
            </sheets><calcPr calcId="152511"/><fileRecoveryPr repairLoad="1"/><extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst></workbook>
            """);

            return stringBuilder.ToString();
        }

        public string GenerateXlStylesXml()
        {
            var stylesBuilder = new StringBuilder();
            stylesBuilder.Append($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
            """);
            var customFormats = _excelCellFormatList.Where(x => x.BuiltCellFormat == BuiltCellFormat.Custom).ToList();
            var customFormatIds = new Dictionary<ExcelCellFormat, int>();
            if (customFormats.Count > 0)
            {
                int customFormatId = 164;
                stylesBuilder.Append($"<numFmts count=\"{customFormats.Count}\">");

                foreach (var customFormat in customFormats)
                {
                    stylesBuilder.Append($"<numFmt numFmtId=\"{customFormatId}\" formatCode=\"{customFormat.CustomFormat}\"/>");
                    customFormatIds.Add(customFormat, customFormatId);
                    customFormatId++;
                }

                stylesBuilder.Append("</numFmts>");
            }

            stylesBuilder.Append($$"""
            <fonts count="1" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><charset val="238"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
            """);

            stylesBuilder.Append("<cellXfs count=\"");
            stylesBuilder.Append(_excelCellFormatList.Count + 1);
            stylesBuilder.Append("\">");
            stylesBuilder.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
            int styleId = 1;

            foreach (var style in _excelCellFormatList)
            {
                switch (style.BuiltCellFormat)
                {
                    case BuiltCellFormat.DateOnly:
                        stylesBuilder.Append("<xf numFmtId=\"14\" applyNumberFormat=\"1\" />");
                        break;
                    case BuiltCellFormat.DateTime:
                        stylesBuilder.Append("<xf numFmtId=\"22\" applyNumberFormat=\"1\" />");
                        break;
                    case BuiltCellFormat.TimeOnly:
                        stylesBuilder.Append("<xf numFmtId=\"21\" applyNumberFormat=\"1\" />");
                        break;
                    case BuiltCellFormat.Custom:
                        stylesBuilder.Append($"<xf numFmtId=\"{customFormatIds[style]}\" applyNumberFormat=\"1\" />");
                        break;
                }

                styleId++;
            }

            stylesBuilder.Append("</cellXfs>");

            stylesBuilder.Append($$"""
            <cellStyles count="1"><cellStyle name="Normalny" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>
            """);

            return stylesBuilder.ToString();
        }

        public string GenerateXlSharedStringsXml() =>
            $$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>
            """;

        public string GenerateXlRels()
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3Styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2Theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
            """);

            var i = 1;
            foreach (var sheet in _workbook.Sheets)
            {
                stringBuilder.Append($"<Relationship Id=\"rId{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>");
                i++;
            }

            stringBuilder.Append($$"""
            <Relationship Id="rId4SharedStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>
            """);

            return stringBuilder.ToString();
        }

        public string GenerateXlTheme() => $$"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Motyw pakietu Office"><a:themeElements><a:clrScheme name="Pakiet Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Pakiet Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Pakiet Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>
        """;

        public void GenerateSheet(Sheet sheet, StreamWriter streamWriter, CancellationToken ct = default)
        {
            streamWriter.Write($$"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
            """);

            streamWriter.Write($$"""
            <sheetViews><sheetView tabSelected="1" workbookViewId="0">
            """);

            if (sheet.View.FreezePanes.XSplit > 0 || sheet.View.FreezePanes.YSplit > 0)
            {
                streamWriter.Write($"<pane {(sheet.View.FreezePanes.YSplit > 0 ? $"ySplit=\"{sheet.View.FreezePanes.YSplit}\"" : "")} {(sheet.View.FreezePanes.XSplit > 0 ? $"xSplit=\"{sheet.View.FreezePanes.XSplit}\"" : "")} topLeftCell=\"{GetAndWriteCellId(sheet.View.FreezePanes.YSplit + 1, sheet.View.FreezePanes.XSplit)}\" activePane=\"bottomLeft\" state=\"frozen\" />");
            }

            streamWriter.Write("</sheetView></sheetViews>");

            streamWriter.Write($$"""
            <sheetFormatPr defaultRowHeight="14.4" x14ac:dyDescent="0.3"/><sheetData>
            """);

            // Get headers
            var recordsType = sheet.Data.GetType().GenericTypeArguments[0];

            var headers = new List<Header>();
            var groupColumnNameInfos = new List<GroupColumnNameInfo>();
            var mergedCells = new List<string>();

            foreach (var property in recordsType.GetProperties())
            {
                int? excelCellFormatId = null;

                if (property.GetCustomAttribute<ColumnFormatAttribute>() != null)
                {
                    excelCellFormatId = ResolveExcelCellFormatId(BuiltCellFormat.Custom, property.GetCustomAttribute<ColumnFormatAttribute>()!._format);
                }
                if (property.PropertyType == typeof(DateTime))
                {
                    excelCellFormatId = ResolveExcelCellFormatId(BuiltCellFormat.DateTime);
                }
                else if (property.PropertyType == typeof(DateOnly))
                {
                    excelCellFormatId = ResolveExcelCellFormatId(BuiltCellFormat.DateOnly);
                }
                else if (property.PropertyType == typeof(TimeOnly))
                {
                    excelCellFormatId = ResolveExcelCellFormatId(BuiltCellFormat.TimeOnly);
                }

                if (property.GetCustomAttribute<ColumnIgnoreAttribute>() == null)
                {
                    headers.Add(new Header
                    {
                        Name = property.GetCustomAttribute<ColumnNameAttribute>()?._name ?? property.Name,
                        Position = property.GetCustomAttribute<ColumnPositionAttribute>()?._index,
                        Property = property,
                        CellFormatId = excelCellFormatId,
                        GroupColumnAttributes = property.GetCustomAttributes<GroupColumnNameAttribute>()
                    });
                }
            }

            var usedPositions = new HashSet<int>(headers.Where(x => x.Position.HasValue).Select(x => x.Position!.Value));
            int rowIndex = 0;
            foreach (var header in headers)
            {
                while (usedPositions.Contains(rowIndex))
                {
                    rowIndex++;
                }

                if (!header.Position.HasValue)
                {
                    header.Position = rowIndex;
                    rowIndex++;
                }
            }

            headers = headers.OrderBy(x => x.Position).ToList();

            int groupColumnIndex = 0;
            // Calculate column groups
            foreach (var header in headers)
            {
                if (header.GroupColumnAttributes is not null)
                {
                    foreach (var groupColumnAttribute in header.GroupColumnAttributes)
                    {
                        bool updated = false;

                        foreach (var groupColumnNameInfo in groupColumnNameInfos)
                        {
                            if (groupColumnNameInfo.Name == groupColumnAttribute._name && groupColumnNameInfo.EndIndex + 1 == groupColumnIndex && groupColumnNameInfo.Depth == groupColumnAttribute._depth)
                            {
                                groupColumnNameInfo.EndIndex++;
                                updated = true;
                            }
                        }

                        if (!updated)
                        {
                            groupColumnNameInfos.Add(
                                new GroupColumnNameInfo
                                {
                                    Name = groupColumnAttribute._name,
                                    Depth = groupColumnAttribute._depth,
                                    StartIndex = groupColumnIndex,
                                    EndIndex = groupColumnIndex
                                });
                        }
                    }
                }

                groupColumnIndex++;
            }

            rowIndex = 1;
            int headerI = 0;

            // Write Column groups
            if (groupColumnNameInfos.Count > 0)
            {
                var rows = new List<(string, List<GroupColumnNameInfo>)>();
                foreach (var group in groupColumnNameInfos.OrderBy(x => x.StartIndex).ThenBy(x => x.Depth).GroupBy(x => x.Depth))
                {
                    streamWriter.Write("<row r=\"");
                    streamWriter.Write(rowIndex.ToString());
                    streamWriter.Write("\" spans=\"1:");
                    streamWriter.Write(headers.Count.ToString());
                    streamWriter.Write("\" x14ac:dyDescent=\"0.3\">");

                    foreach (var groupColumnNameInfo in group)
                    {
                        //<c r="A1"><v>11</v></c>
                        streamWriter.Write("<c r=\"");
                        streamWriter.Write(GetAndWriteCellId(rowIndex, groupColumnNameInfo.StartIndex));
                        streamWriter.Write("\" t=\"inlineStr\">");
                        streamWriter.Write("<is><t>");
                        streamWriter.Write(groupColumnNameInfo.Name);
                        streamWriter.Write("</t></is>");
                        streamWriter.Write("</c>");

                        mergedCells.Add($"{GetAndWriteCellId(rowIndex, groupColumnNameInfo.StartIndex)}:{GetAndWriteCellId(rowIndex, groupColumnNameInfo.EndIndex)}");
                    }

                    rowIndex++;

                    streamWriter.Write("</row>");
                }
            }

            // Write header
            // <row r="1" spans="1:3" x14ac:dyDescent="0.3"><c r="A1"><v>11</v></c></row>
            streamWriter.Write("<row r=\"");
            streamWriter.Write(rowIndex.ToString());
            streamWriter.Write("\" spans=\"1:");
            streamWriter.Write(headers.Count.ToString());
            streamWriter.Write("\" x14ac:dyDescent=\"0.3\">");
            headerI = 0;

            foreach (var header in headers)
            {
                //<c r="A1"><v>11</v></c>
                streamWriter.Write("<c r=\"");
                streamWriter.Write(GetAndWriteCellId(rowIndex, headerI));
                streamWriter.Write("\" t=\"inlineStr\">");
                streamWriter.Write("<is><t>");
                streamWriter.Write(header.Name);
                streamWriter.Write("</t></is>");
                streamWriter.Write("</c>");

                headerI++;
            }

            rowIndex++;

            streamWriter.Write("</row>");

            // TODO!
            foreach (var data in sheet.Data)
            {
                // <c r="A1"><v>11</v></c>   - numbers, data, strings when used shared string table
                // <c r="A1"><is><t>string</t></is></c> - inline strings

                streamWriter.Write("<row r=\"");
                streamWriter.Write(rowIndex.ToString());
                streamWriter.Write("\" spans=\"1:");
                streamWriter.Write(headers.Count.ToString());
                streamWriter.Write("\" x14ac:dyDescent=\"0.3\">");

                int columnIndex = 0;
                foreach (var header in headers)
                {
                    var value = header.Property?.GetValue(data);
                    if (value is not null)
                    {
                        //<c r="A1"><v>11</v></c>
                        streamWriter.Write("<c r=\"");
                        GetAndWriteCellId(rowIndex, columnIndex);
                        streamWriter.Write("\" ");

                        if (header.CellFormatId.HasValue)
                        {
                            streamWriter.Write("s=\"");
                            streamWriter.Write(header.CellFormatId);
                            streamWriter.Write("\" ");
                        }

                        if (value is string stringValue)
                        {
                            streamWriter.Write("t=\"inlineStr\"><is><t>");
                            streamWriter.Write(stringValue);
                            streamWriter.Write("</t></is>");
                        }
                        else if (value is int intValue)
                        {
                            streamWriter.Write("><v>");
                            streamWriter.Write(intValue);
                            streamWriter.Write("</v>");
                        }
                        else if (value is long longValue)
                        {
                            streamWriter.Write("><v>");
                            streamWriter.Write(longValue);
                            streamWriter.Write("</v>");
                        }
                        else if (value is float floatValue)
                        {
                            streamWriter.Write("><v>");
                            streamWriter.Write(floatValue.ToString(new CultureInfo("en-US")));
                            streamWriter.Write("</v>");
                        }
                        else if (value is double doubleValue)
                        {
                            streamWriter.Write("><v>");
                            streamWriter.Write(doubleValue.ToString(new CultureInfo("en-US")));
                            streamWriter.Write("</v>");
                        }
                        else if (value is bool boolValue)
                        {
                            streamWriter.Write("t=\"b\"><v>");
                            streamWriter.Write(boolValue ? 1 : 0);
                            streamWriter.Write("</v>");
                        }
                        else if (value is DateTime dateTimeValue)
                        {
                            streamWriter.Write("><v t=\"n\">");
                            streamWriter.Write(DateTimeToDouble(dateTimeValue).ToString(new CultureInfo("en-US")));
                            streamWriter.Write("</v>");
                        }
                        else if (value is DateOnly dateOnlyValue)
                        {
                            streamWriter.Write("><v t=\"n\">");
                            streamWriter.Write(DateTimeToDouble(dateOnlyValue).ToString(new CultureInfo("en-US")));
                            streamWriter.Write("</v>");
                        }
                        else if (value is TimeOnly timeOnlyValue)
                        {
                            streamWriter.Write("><v t=\"n\">");
                            streamWriter.Write(DateTimeToDouble(timeOnlyValue).ToString(new CultureInfo("en-US")));
                            streamWriter.Write("</v>");
                        }
                        else
                        {
                            streamWriter.Write("t=\"inlineStr\"><is><t>");
                            streamWriter.Write(value?.ToString());
                            streamWriter.Write("</t></is>");
                        }

                        streamWriter.Write("</c>");
                    }

                    columnIndex++;
                }

                rowIndex++;

                streamWriter.Write("</row>");

                ct.ThrowIfCancellationRequested();
            }

            streamWriter.Write("</sheetData>");

            if (mergedCells.Count > 0)
            {
                streamWriter.Write($"<mergeCells count=\"{mergedCells.Count}\">");

                foreach (var range in mergedCells)
                {
                    streamWriter.Write($"<mergeCell ref=\"{range}\" />");
                }

                streamWriter.Write("</mergeCells>");
            }

            streamWriter.Write($$"""
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>
            """);
        }

        private int? ResolveExcelCellFormatId(BuiltCellFormat builtCellFormat, string? customFormat = null)
        {
            int? excelCellFormatId;
            var excelCellFormat = _excelCellFormatList.FirstOrDefault(x => (builtCellFormat != BuiltCellFormat.Custom && x.BuiltCellFormat == builtCellFormat) || (builtCellFormat == BuiltCellFormat.Custom && x.CustomFormat == customFormat));
            if (excelCellFormat == null)
            {
                excelCellFormat = new ExcelCellFormat() { BuiltCellFormat = builtCellFormat, CustomFormat = customFormat };
                _excelCellFormatList.Add(excelCellFormat);
            }

            excelCellFormatId = _excelCellFormatList.IndexOf(excelCellFormat) + 1; //0 is default style
            return excelCellFormatId;
        }

        public string GetAndWriteCellId(int rowIndex, int columnIndex)
        {
            var stringBuilder = new StringBuilder(13);
            stringBuilder.Append(GetColumnName(columnIndex));
            stringBuilder.Append(rowIndex);
            return stringBuilder.ToString();
        }

        public void GetCellId(StreamWriter streamWriter, int rowIndex, int columnIndex)
        {
            streamWriter.Write(GetColumnName(columnIndex));
            streamWriter.Write(rowIndex);
        }

        public string GetColumnName(int columnIndex)
        {
            int remainder = 0;
            Span<char> columnNameChars = stackalloc char[3];
            int byteIndex = 2;
            columnIndex++;

            while (columnIndex > 0)
            {
                remainder = (columnIndex - 1) % 26;
                columnNameChars[byteIndex--] = (char)(remainder + 'A');
                columnIndex = (columnIndex - 1) / 26;
            }

            return new string(columnNameChars[(byteIndex + 1)..]);
        }

        private double DateTimeToDouble(DateTime dateTime)
        {
            var startDate = new DateTime(1900, 01, 01, 0, 0, 0, kind: dateTime.Kind);
            var totalDays = (dateTime - startDate).TotalDays + 1; // 1900-01-01 is 1 day
            if (dateTime > new DateTime(1900, 02, 28, 23, 59, 59, 999)) // 1900-02-29 is for excel existing day
            {
                totalDays++;
            }

            return totalDays;
        }

        private double DateTimeToDouble(DateOnly dateOnly)
        {
            var dateTime = dateOnly.ToDateTime(new TimeOnly(0, 0, 0), DateTimeKind.Utc);
            var startDate = new DateTime(1900, 01, 01, 0, 0, 0, kind: DateTimeKind.Utc);
            var totalDays = (dateTime - startDate).TotalDays + 1; // 1900-01-01 is 1 day
            if (dateTime > new DateTime(1900, 02, 28, 23, 59, 59, 999)) // 1900-02-29 is for excel existing day
            {
                totalDays++;
            }

            return totalDays;
        }

        private double DateTimeToDouble(TimeOnly timeOnly)
        {
            var totalMilliseconds = (timeOnly - new TimeOnly(0, 0, 0, 0)).TotalMilliseconds;
            return totalMilliseconds / 86399999;
        }

        private class Header
        {
            public string? Name { get; set; }
            public PropertyInfo? Property { get; set; }
            public int? Position { get; set; }
            public int? CellFormatId { get; set; }
            public IEnumerable<GroupColumnNameAttribute>? GroupColumnAttributes { get; set; } = new List<GroupColumnNameAttribute>();
        }

        private class GroupColumnNameInfo
        {
            public string? Name { get; set; }

            public int StartIndex { get; set; }

            public int EndIndex { get; set; }

            public int Depth { get; set; }
        }

        private class ExcelCellFormat
        {
            public BuiltCellFormat BuiltCellFormat { get; set; }

            public string? CustomFormat { get; set; }
        }

        private enum BuiltCellFormat
        {
            DateTime = 1,
            DateOnly = 2,
            TimeOnly = 3,
            Custom = 4,
        }
    }
}
