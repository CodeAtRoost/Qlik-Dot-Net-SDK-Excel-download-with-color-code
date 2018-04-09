using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Qlik.Engine;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Web.Hosting;
using System.Drawing;

namespace QlikService.Models
{
    

    public class excelDownload
    {
        private static string return_name;
        public static string createExcel(string appId, string objectid)
        {
            var location = Qlik.Engine.Location.FromUri(new Uri("ws://localhost:4848"));
            try
            {
                // Connect to desktop

                location.AsDirectConnectionToPersonalEdition();
            }
            catch (SystemException e)
            {
                Console.WriteLine("Could not open app! " + e.ToString());
                return   "Connection to Qlik engine failed";

            }
            //  location.AsNtlmUserViaProxy(proxyUsesSsl: false);
            try
            {
                // Open the app with name "Beginner's tutorial"
                var appIdentifier = location.AppWithNameOrDefault(@"Consumer Sales", noVersionCheck: true);
                using (var app = location.App(appIdentifier, noVersionCheck: true))
                {
                    // Clear all selections to set the app in an known state
                    // app.ClearAll();
                    var  viz=app.GetGenericObject("akDGX");
                    
                    Qlik.Sense.Client.Visualizations.Table obj = app.GetObject<Qlik.Sense.Client.Visualizations.Table>(objectid);
                    // Get the sheet with the title "Dashboard"
                
                    var first10CellsPage = new NxPage { Top = 0, Left = 0, Width = 5, Height = 1000 };
                    IEnumerable<NxDataPage> data = obj.GetHyperCubeData("/qHyperCubeDef", new[] { first10CellsPage });
                    
                    var minfo = obj.MeasureInfo;
                    int dimcount = obj.DimensionInfo.Count();
                    int measurecount = obj.MeasureInfo.Count();
                    int rowcount = data.ElementAt(0).Matrix.Count();
                    Console.WriteLine(obj.DimensionInfo.Count());
                    Console.WriteLine(obj.MeasureInfo.Count());
                    Console.WriteLine(data.ElementAt(0).Matrix.Count());
                    
                    string[,] header = new string[data.ElementAt(0).Matrix.Count() + 1,dimcount+measurecount];
                    double[,] data_color = new double[data.ElementAt(0).Matrix.Count() + 1, dimcount + measurecount];
                    string[,] dim_alternate_title = new string[dimcount, 2];
                    for (int i = 0; i <dimcount; i++)
                    {
                        header[0, i ] = obj.DimensionInfo.ElementAt(i).FallbackTitle;
                        dim_alternate_title[i, 0] = obj.DimensionInfo.ElementAt(i).OthersLabel;
                   
                    }
                    for (int i = 0; i < measurecount; i++)
                    {
                        header[0, i + dimcount] = obj.MeasureInfo.ElementAt(i).FallbackTitle;
                   
                    }
                    for (int i = 0; i < rowcount; i++)
                    {
                        for (int j = 0; j < data.ElementAt(0).Matrix.ElementAt(0).Count(); j++)
                        {
                            NxCell el = data.ElementAt(0).Matrix.ElementAt(i).ElementAt(j);
                            double col=0;
                            if (el.AttrExps!= null )  col= el.AttrExps.Values.ElementAt (0).Num.ToString ()!="NaN"? el.AttrExps.Values.ElementAt(0).Num:0;
                            if (data.ElementAt(0).Matrix.ElementAt(0).ElementAt(j).IsOtherCell)
                            {
                                header[1 + i, j] = dim_alternate_title[i,0];
                                data_color[1 + i, j] = col;
                            }
                            else
                            {
                                header[1 + i, j] = data.ElementAt(0).Matrix.ElementAt(i).ElementAt(j).Text;
                                data_color[1 + i, j] = col;
                            }
                            
                            if (j >= dimcount) {
                            }

                        }

                    }
                    using (ExcelPackage excel = new ExcelPackage())
                    {
                        excel.Workbook.Worksheets.Add("Worksheet1");
                        var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                        
                        for (int i = 0; i < header.GetLength(0); i++)
                        {

                           for (int j = 0; j < header.GetLength(1); j++)
                            {
                                ExcelRange range = worksheet.Cells[i + 1, j + 1];
                                range.Value = header[i, j];
                                 if (data_color[i, j] != 0)
                                {
                                    byte[] bytes = BitConverter.GetBytes(data_color[i,j]);
                                    string hex=DoubleToHex(data_color[i, j], 0);
                                    Color bg = System.Drawing.ColorTranslator.FromHtml("#" + hex);
                                    try
                                    {
                                        worksheet.Cells[i + 1, j + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        worksheet.Cells[i + 1, j + 1].Style.Fill.BackgroundColor.SetColor(bg);
                                    }
                                    catch (Exception e) { return e.Message + "\n"+   e.StackTrace.ToString(); }
                                }
                   
                            }

                        }
                        FileInfo excelFile = new FileInfo(HostingEnvironment.ApplicationPhysicalPath + @"\temp\"+"test.xlsx");
                        excel.SaveAs(excelFile);
                        return_name= HostingEnvironment.ApplicationVirtualPath+ @"temp/" + "test.xlsx";
                        

                    }
                    return return_name;
                }
            }
            catch (SystemException e)
            {
                Console.WriteLine("Could not open app! " + e.ToString());
                return "Could not open the app";
        }

    }

       internal static string DoubleToHex(double value, int maxDecimals)
        {
            string result = string.Empty;
            if (value < 0)
            {
                result += "-";
                value = -value;
            }
            if (value > ulong.MaxValue)
            {
                result += double.PositiveInfinity.ToString();
                return result;
            }
            ulong trunc = (ulong)value;
            result += trunc.ToString("X");
            value -= trunc;
            if (value == 0)
            {
                return result;
            }
            result += ".";
            byte hexdigit;
            while ((value != 0) && (maxDecimals != 0))
            {
                value *= 16;
                hexdigit = (byte)value;
                result += hexdigit.ToString("X");
                value -= hexdigit;
                maxDecimals--;
            }
            return result;
        }

        internal static ulong HexToUInt64(string hex)
        {
            ulong result;
            if (ulong.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            throw new ArgumentException("Cannot parse hex string.", "hex");
        }



    }

        

       

        
}
