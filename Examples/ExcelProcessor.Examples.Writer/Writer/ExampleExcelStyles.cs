using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Core;

namespace ExcelProcessor.Example.Writer
{
    internal class ExampleExcelStyles : ExcelStyles
    {
        public const string RedCell = "RedCell";
        public const string GreenArialCell = "GreenArialCell";
        public const string HeaderTableCell = "HeaderTableCell";
        public const string TableCell = "TableCell";
        public const string BlueCell = "BlueCell";
        public const string LightBlueCell = "LightBlueCell";
        public const string Orange = "Orange";
        public const string LightOrange = "LightOrange";

        public override void Inyect(Stylesheet stylesSheet)
        {
            uint fullBlackBorder = InyectBorder(stylesSheet, GenerateBorder("000000"));

            Font arial13Font = new Font();
            arial13Font.Append(new FontSize() { Val = 13D });
            arial13Font.Append(new FontName() { Val = "Arial" });

            uint arialFont = InyectFont(stylesSheet, arial13Font);

            InyectStyle(RedCell, stylesSheet,
                        new CellFormat()
                        {
                            FillId = InyectFill(stylesSheet, GenerateSolidFill("FF9191")),
                            Alignment = new Alignment()
                            {
                                WrapText = true,
                                Vertical = VerticalAlignmentValues.Center,
                                Horizontal = HorizontalAlignmentValues.Left
                            },
                            BorderId = fullBlackBorder,
                        });

            InyectStyle(GreenArialCell, stylesSheet,
                        new CellFormat()
                        {
                            FontId = arialFont,
                            FillId = InyectFill(stylesSheet, GenerateSolidFill("009821")),
                            Alignment = new Alignment()
                            {
                                WrapText = true,
                                Vertical = VerticalAlignmentValues.Center,
                                Horizontal = HorizontalAlignmentValues.Left
                            },
                            BorderId = fullBlackBorder,
                        });

            InyectStyle(HeaderTableCell, stylesSheet,
                        new CellFormat()
                        {
                            FillId = InyectFill(stylesSheet, GenerateSolidFill("A0A0A0")),
                            Alignment = new Alignment()
                            {
                                WrapText = true,
                                Vertical = VerticalAlignmentValues.Center,
                                Horizontal = HorizontalAlignmentValues.Center
                            },
                            BorderId = fullBlackBorder,
                        });

            InyectStyle(TableCell, stylesSheet,
                        new CellFormat()
                        {
                            FillId = InyectFill(stylesSheet, GenerateSolidFill("CCCCCC")),
                            Alignment = new Alignment()
                            {
                                WrapText = true,
                                Vertical = VerticalAlignmentValues.Center,
                                Horizontal = HorizontalAlignmentValues.Center
                            },
                            BorderId = fullBlackBorder,
                        });

            InyectStyle(BlueCell, stylesSheet,
                        new CellFormat()
                        {
                            FillId = InyectFill(stylesSheet, GenerateSolidFill("389AFF")),
                            Alignment = new Alignment()
                            {
                                WrapText = true,
                                Vertical = VerticalAlignmentValues.Center,
                                Horizontal = HorizontalAlignmentValues.Center
                            },
                            BorderId = fullBlackBorder,
                        });

            InyectStyle(LightBlueCell, stylesSheet,
                        new CellFormat()
                        {
                            FillId = InyectFill(stylesSheet, GenerateSolidFill("6BC5FF")),
                            Alignment = new Alignment()
                            {
                                WrapText = true,
                                Vertical = VerticalAlignmentValues.Center,
                                Horizontal = HorizontalAlignmentValues.Center
                            },
                            BorderId = fullBlackBorder,
                        });

            InyectStyle(Orange, stylesSheet,
                        new CellFormat()
                        {
                            FillId = InyectFill(stylesSheet, GenerateSolidFill("F8CBAD")),
                            Alignment = new Alignment()
                            {
                                WrapText = true,
                                Vertical = VerticalAlignmentValues.Center,
                                Horizontal = HorizontalAlignmentValues.Center
                            },
                            BorderId = fullBlackBorder,
                        });

            InyectStyle(LightOrange, stylesSheet,
                        new CellFormat()
                        {
                            FillId = InyectFill(stylesSheet, GenerateSolidFill("FCE7D8")),
                            Alignment = new Alignment()
                            {
                                WrapText = true,
                                Vertical = VerticalAlignmentValues.Center,
                                Horizontal = HorizontalAlignmentValues.Center
                            },
                            BorderId = fullBlackBorder,
                        });

        }
    }
}
