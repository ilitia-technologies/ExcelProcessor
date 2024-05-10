using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProcessor.Abstractions;

namespace ExcelProcessor.Core
{
    public abstract class ExcelStyles : IExcelStyles
    {
        protected readonly Dictionary<string, uint> styles = new Dictionary<string, uint>();

        public virtual void Inyect(Stylesheet stylesSheet)
        {

        }

        protected void InyectStyle(string name, Stylesheet stylesSheet, CellFormat cellFormat)
        {
            stylesSheet.CellFormats.Append(cellFormat);
            styles.Add(name, (uint)stylesSheet.CellFormats.Elements<CellFormat>().Count() - 1);
        }

        protected uint InyectFill(Stylesheet stylesSheet, Fill fill)
        {
            stylesSheet.Fills.AppendChild(fill);
            return (uint)stylesSheet.Fills.Elements<Fill>().Count() - 1;
        }

        protected uint InyectBorder(Stylesheet stylesSheet, Border fill)
        {
            stylesSheet.Borders.AppendChild(fill);
            return (uint)stylesSheet.Borders.Elements<Border>().Count() - 1;
        }

        protected uint InyectFont(Stylesheet stylesSheet, Font font)
        {
            stylesSheet.Fonts.AppendChild(font);
            return (uint)stylesSheet.Fonts.Elements<Font>().Count() - 1;
        }

        public uint GetStyleId(string styleName)
        {
            if (!styles.ContainsKey(styleName))
                throw new ArgumentException($"{styleName} not found", nameof(styleName));

            return styles[styleName];
        }

        protected Fill GenerateSolidFill(HexBinaryValue color)
        {
            return new Fill()
            {
                PatternFill = new PatternFill()
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor() { Rgb = color }
                }
            };
        }

        protected Border GenerateBorder(string colorAsRgbString)
        {            
            LeftBorder leftBorder = new LeftBorder() { Style = BorderStyleValues.Thin };
            leftBorder.Append(new Color() { Rgb = colorAsRgbString });

            RightBorder rightBorder = new RightBorder() { Style = BorderStyleValues.Thin };
            rightBorder.Append(new Color() { Rgb = colorAsRgbString });

            TopBorder topBorder = new TopBorder() { Style = BorderStyleValues.Thin };
            topBorder.Append(new Color() { Rgb = colorAsRgbString });

            BottomBorder bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin };
            bottomBorder.Append(new Color() { Rgb = colorAsRgbString });

            Border border = new Border();
            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            return border;

        }
    }
}
