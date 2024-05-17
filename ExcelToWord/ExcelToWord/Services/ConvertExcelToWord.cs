using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelToWord.Model;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using NPOI.HPSF;
using System.Reflection;

namespace ExcelToWord.Services
{
    static class ConvertExcelToWord
    {
        public static List<Person> People = new();

        public static string DebtString = " - Id számú bizonylat ellenértékét (Debt Ft – KindOfDebt, fizetési határidő lejárta: Date)";
        
        public static void ConvertFile(string excel, string word, string folder)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage xlPackage = new(new FileInfo(excel)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                Debug.WriteLine(totalRows);

                var sb = new StringBuilder(); //this is your data
                for (int rowNum = 4; rowNum <= totalRows; rowNum++) //select starting row here
                {
                    if (myWorksheet.GetValue(rowNum, 1) != null) { 
                    if (!People.Any(x => x.Name == myWorksheet.GetValue(rowNum, 1).ToString()))
                    {
                        Debug.WriteLine(myWorksheet.GetValue(rowNum, 1).ToString());
                        var Address = myWorksheet.GetValue(rowNum, 15).ToString();
                        Address = Address.Remove(0,14);
                        var splitted = Address.Split(",");

                        People.Add(new Person(myWorksheet.GetValue(rowNum, 1).ToString(),
                            (double)myWorksheet.GetValue(rowNum, 9),
                            (double)myWorksheet.GetValue(rowNum, 14),
                            splitted[0],
                            splitted[1],
                            new Debt((double)myWorksheet.GetValue(rowNum, 2),
                                    myWorksheet.GetValue(rowNum, 3).ToString(), 
                                    (double)myWorksheet.GetValue(rowNum, 6), 
                                    myWorksheet.GetValue(rowNum, 22).ToString())));
                    }
                    else
                    {
                        People.Find(x => x.Name == myWorksheet.GetValue(rowNum, 1).ToString()).Debts.Add(new Debt((double)myWorksheet.GetValue(rowNum, 2), 
                            myWorksheet.GetValue(rowNum, 3).ToString(), 
                            (double)myWorksheet.GetValue(rowNum, 6), 
                            myWorksheet.GetValue(rowNum, 22).ToString()));
                    }
                    }
                }
            }

            foreach (var person in People) {
                using (var originalDocument = WordprocessingDocument.Open(word, false))
                using (var document = WordprocessingDocument.Create($@"{folder}\{person.Name}.docx", WordprocessingDocumentType.Document))
                {
                    foreach (var part in originalDocument.Parts)
                        document.AddPart(part.OpenXmlPart, part.RelationshipId);

                    // Gets the MainDocumentPart of the WordprocessingDocument 
                    var main = document.MainDocumentPart;
                    // document fonts
                    var fonts = main.FontTablePart;
                    // document styles
                    var styles = main.StyleDefinitionsPart;
                    var effects = main.StylesWithEffectsPart;
                    // root element part of the doc
                    var doc = main.Document;
                    // actual document body
                    var body = doc.Body;

                    var paras = body.Elements<Paragraph>();

                    List<string> DebtsStrings = new();

                    foreach (var debt in person.Debts)
                    {
                        var substituted = DebtString;
                        substituted = substituted.Replace("Id", debt.Id.ToString());
                        substituted = substituted.Replace("KindOfDebt", debt.Name);
                        substituted = substituted.Replace("Debt", debt.Amount.ToString("0,0"));
                        substituted = substituted.Replace("Date", debt.Deadline.ToString("yyyy.MM.dd."));
                        DebtsStrings.Add(substituted);
                    }
                    /*
                    var fullDebtString = "";

                    var debtPara = new Paragraph(new Run());
                    foreach (var run in debtPara)
                    {
                        foreach (var debtstring in DebtsStrings)
                        {
                            run.Append(new Text(debtstring));
                            run.Append(new Break());
                        }
                    }
                    */
                    foreach (var para in paras)
                    {
                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (var text in run.Elements<Text>())
                            {
                                if (text.Text.Contains("Item"))
                                {
                                    for (var i = 0; i < DebtsStrings.Count(); i++)
                                    {
                                        run.AppendChild<Paragraph>(new Paragraph(new Text(DebtsStrings[i])));
                                    }
                                }
                            }
                        }
                    }

                    foreach (var para in paras)
                    {
                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (var text in run.Elements<Text>())
                            {
                                 if (text.Text.Contains("Item"))
                                {
                                    text.Text = text.Text.Replace("Item", "");
                                }
                                 else if (text.Text.Contains("Name"))
                                {
                                    text.Text = text.Text.Replace("Name", person.Name);
                                }
                                else if (text.Text.Contains("AllDebtWithInterest"))
                                {
                                    text.Text = text.Text.Replace("AllDebtWithInterest", (person.AllDebt + person.AllInterest).ToString("0,0"));
                                }
                                else if (text.Text.Contains("City"))
                                {
                                    text.Text = text.Text.Replace("City", person.City);
                                }
                                else if (text.Text.Contains("Street"))
                                {
                                    text.Text = text.Text.Replace("Street", person.Street);
                                }
                                else if (text.Text.Contains("AllDebt"))
                                {
                                    text.Text = text.Text.Replace("AllDebt", person.AllDebt.ToString("0,0"));
                                }
                                else if (text.Text.Contains("Interest"))
                                {
                                    text.Text = text.Text.Replace("Interest", person.AllInterest.ToString("0,0"));
                                }
                                else if (text.Text.Contains("DateNow"))
                                {
                                    text.Text = text.Text.Replace("DateNow", DateTime.Now.ToString("yyyy-MM-dd"));
                                }
                                
                            }
                        }
                    }

                }
            }
            
        }
    }
}
