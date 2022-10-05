using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelGeneratingApp.Models
{
    public class ExcelLib
    {
        // Default Constructor.
        public ExcelLib() { }
        // Method for adding Header and Title of the table containing List of object values.
        public IXLWorksheet addHeader(XLWorkbook wb, List<Object> objs, string title,  string fontFamily = "Sakkal Majalla", string color = "#3498DB")
        {
            // Sheet initialisation.
            var ws = wb.Worksheets.Add("nomDeLaListe").SetTabColor(XLColor.UaBlue);
            // font choice.
            ws.Style.Font.FontName = fontFamily;
            ws.Style.Font.SetFontSize(13);
            ws.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Style.Alignment.WrapText = true;
            Object obj = objs.FirstOrDefault();
            // Add the model fields to the header of the excel file.
            int totalOfFields = obj.GetType().GetProperties().Length; // number of fields in the object.
            int numberOfFields = 0;
            // Adding the title of table in excel file.
            ws.Range(ws.Cell(4, 4), ws.Cell(4, totalOfFields + 3)).Merge().Value = title;
            ws.Range(ws.Cell(4, 4), ws.Cell(4, totalOfFields + 3)).Merge().Style.Fill.BackgroundColor = XLColor.FromHtml(color); ;
            ws.Range(ws.Cell(4, 4), ws.Cell(4, totalOfFields + 3)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(ws.Cell(4, 4), ws.Cell(4, totalOfFields + 3)).Style.Font.Bold = true;
            ws.Range(ws.Cell(4, 4), ws.Cell(4, totalOfFields + 3)).Style.Font.FontColor = XLColor.WhiteSmoke;
            ws.Range(ws.Cell(4, 4), ws.Cell(4, totalOfFields + 3)).Style.Font.FontSize = 18;
            //Looping all propeties of the object.
            foreach (var prop in obj.GetType().GetProperties())
            {   
                    var displayNameAttribute = prop.GetCustomAttributes(typeof(DisplayNameAttribute), false);
                    string displayName = prop.Name;
                    if (displayNameAttribute.Count() != 0)
                    {
                        displayName = (displayNameAttribute[0] as DisplayNameAttribute).DisplayName;
                    }
                    numberOfFields++;
                    ws.Cell(5, totalOfFields - numberOfFields + 4).Value = displayName;
                    ws.Cell(5, totalOfFields - numberOfFields + 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    ws.Cell(5, totalOfFields - numberOfFields + 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    ws.Cell(5, totalOfFields - numberOfFields + 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    ws.Cell(5, totalOfFields - numberOfFields + 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    ws.Column(totalOfFields - numberOfFields + 4).Width = 30;
                    ws.Column(totalOfFields - numberOfFields + 4).Style.Font.Bold = true;
            }
            ws.Range(ws.Cell(5, 4), ws.Cell(5, totalOfFields + 3)).SetAutoFilter();
            return ws;
        }
        public IXLWorksheet addBody(IXLWorksheet ws, List<Object> objs)
        {
            int numberOfFields = 0;
            int numberOfRecords = 0;
            Object obj = objs.FirstOrDefault();
            int totalOfFields = obj.GetType().GetProperties().Length;
            string previousValue = "";
            int indexOfPreviousValue = 0;

            foreach (var item in objs.ToList())
            {
                numberOfFields = 0;
                Type myType = item.GetType();
                IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());

                foreach (PropertyInfo prop in props)
                {
                    object propValue = prop.GetValue(item, null);

                    numberOfFields++;
                    ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4).Value = propValue;

                    ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4).Style.Font.Bold = true;

                    ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    if (numberOfFields == 1 && numberOfRecords == 0)
                    {
                        previousValue = propValue.ToString();
                    }
                    else
                    {
                        if (numberOfFields == 1)
                        {
                            if (previousValue == propValue.ToString())
                            {
                                ws.Range(ws.Cell(6 + numberOfRecords - (1 + indexOfPreviousValue), totalOfFields - numberOfFields + 4), ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4)).Merge().Value = propValue.ToString();

                                ws.Range(ws.Cell(6 + numberOfRecords - (1 + indexOfPreviousValue), totalOfFields - numberOfFields + 4), ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4)).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                ws.Range(ws.Cell(6 + numberOfRecords - (1 + indexOfPreviousValue), totalOfFields - numberOfFields + 4), ws.Cell(6 + numberOfRecords, totalOfFields - numberOfFields + 4)).Merge().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                                indexOfPreviousValue++;
                            }
                            else
                            {
                                previousValue = propValue.ToString();
                                indexOfPreviousValue = 0;
                            }
                        }

                    }


                }
                numberOfRecords++;
            }

            return ws;
        }
        public void Generate(List<Object> objs, string title, string fontFamily = "Sakkal Majalla", string color = "#3498DB")
        {
            // Workbook creation.
            using (XLWorkbook wb = new XLWorkbook())
            {
                    var ws = addHeader(wb, objs, title);
                        ws = addBody(ws, objs);
                        wb.SaveAs("C://TestExcelGen.xlsx");
            }
        }

    }

}