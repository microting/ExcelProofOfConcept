using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microting.eFormApi.BasePn.Abstractions;
using Microting.eFormApi.BasePn.Infrastructure.Models.API;
using Microting.eFormApi.BasePn.Infrastructure.Helpers.PluginDbOptions;
using System.Reflection;
using System.IO;
using Castle.Core.Internal;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using eFormCore;
using Microsoft.EntityFrameworkCore.Internal;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using Microting.eForm.Dto;
using Microting.eForm.Infrastructure;
using Microting.eForm.Infrastructure.Data.Entities;
using Microting.eForm.Infrastructure.Models;
using Microting.eFormApi.BasePn.Infrastructure.Helpers;
using OpenStack.NetCoreSwiftClient.Extensions;
using Constants = Microting.eForm.Infrastructure.Constants.Constants;
using OfficeOpenXml;
using RabbitMQ.Client.Logging;
using KeyValuePair = Microting.eForm.Dto.KeyValuePair;

namespace ExcelReadManipPOC
{
    public class Class1
    {
        private static Core _core;
        
        
        public static async Task Main()
        {
            
            _core = new Core();

            var connection = _core
                .StartSqlOnly(
                    "host= localhost;Database=420_SDK;user = root;port=3306;Convert Zero Datetime = true;SslMode=none;")
                .Result;

            MicrotingDbContext dbContext = _core.dbContextHelper.GetDbContext();

            question_sets questionSets = new question_sets
            {
                Name = "Test-Set"
            };
            if (dbContext.question_sets.Count(x => x.Name == questionSets.Name) != 1)
            {
               await questionSets.Create(dbContext);    
            }
            
            languages language = new languages
            {
                Description = "Description",
                Name = "da-DK"
            };
            if (dbContext.languages.Count(x => x.Name == "da-DK") != 1)
            {
                await language.Create(dbContext);
            }

            languages dbLanguage = await dbContext.languages.FirstOrDefaultAsync(x => x.Name == language.Name);

            question_sets dbQuestionSets = await dbContext.question_sets.FirstOrDefaultAsync(x => x.Name == questionSets.Name);

            string[] questionNames = new[] {"Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8", "Q9", "Q10", "Q11", "Q12", "Q13"};
            //                                13    14    15    16    17    18    19    20    21    22    23    24
            List<KeyValuePair<int, questions>> questionIds = new List<KeyValuePair<int, questions>>();

            int qi = 13;
            foreach (var questionName in questionNames)
            {
                if (questionName != "Q13" && questionName != "Q1")
                {
                    var questionTranslation = 
                        dbContext.QuestionTranslations.SingleOrDefault(x => x.Name == questionName);

                    if (questionTranslation == null)
                    {
                        questions question = new questions()
                        {
                            QuestionSetId = dbQuestionSets.Id,
                            QuestionType = Constants.QuestionTypes.Smiley2
                        };
                        await question.Create(dbContext);

                        KeyValuePair<int, questions> kvp = new KeyValuePair<int, questions>(qi, question);
                        questionIds.Add(kvp);
                    
                        questionTranslation = new question_translations()
                        {
                            Name = questionName,
                            QuestionId = question.Id,
                            LanguageId = dbLanguage.Id
                        };
                        await questionTranslation.Create(dbContext);
                    }
                    else
                    {
                        KeyValuePair<int, questions> kvp = new KeyValuePair<int, questions>(qi, questionTranslation.Question);
                        questionIds.Add(kvp);
                    }
                }
                else
                {
                    var questionTranslation = 
                        dbContext.QuestionTranslations.SingleOrDefault(x => x.Name == questionName);

                    if (questionTranslation == null)
                    {
                        questions question = new questions()
                        {
                            QuestionSetId = dbQuestionSets.Id,
                            QuestionType = questionName == "Q1" ? Constants.QuestionTypes.List : Constants.QuestionTypes.Multi
                        };
                        await question.Create(dbContext);

                        questionTranslation = new question_translations()
                        {
                            Name = questionName,
                            QuestionId = question.Id,
                            LanguageId = dbLanguage.Id
                        };
                        await questionTranslation.Create(dbContext);

                        string[] questionOptions;
                        if (questionName == "Q1")
                        {
                            questionOptions = new[] {"Ja", "Nej"};    
                        }
                        else
                        {
                            questionOptions = new[] {"1", "2", "3", "4", "5"};
                        }
                        

                        foreach (string questionOption in questionOptions)
                        {
                            options option = new options()
                            {
                                QuestionId = question.Id,
                                Weight = 1,
                                WeightValue = 1
                            };
                            await option.Create(dbContext);
                            
                            option_translations optionTranslation = new option_translations()
                            {
                                OptionId = option.Id,
                                Name = questionOption,
                                LanguageId = dbLanguage.Id
                            };

                            await optionTranslation.Create(dbContext);
                        }
                        
                        KeyValuePair<int, questions> kvp = new KeyValuePair<int, questions>(qi, question);
                        questionIds.Add(kvp);
                    }
                    else
                    {
                        KeyValuePair<int, questions> kvp = new KeyValuePair<int, questions>(qi, questionTranslation.Question);
                        questionIds.Add(kvp);
                    }
                }
                qi++;
            }
            
            // Q13 with options
            
            // KeyValuePair<int, questions> kvp = new KeyValuePair<int, questions>(qi, questionTranslation.Question);
            // questionIds.Add(kvp);

            survey_configurations surveyConfiguration = new survey_configurations
            {
                QuestionSetId = dbQuestionSets.Id,
                Name = "Configuartion 1"
            };
            if (dbContext.survey_configurations.Count(x => x.Name == surveyConfiguration.Name) != 1)
            {
                await surveyConfiguration.Create(dbContext);
            }

            survey_configurations dbSurveyConfiguration =
                await dbContext.survey_configurations.FirstOrDefaultAsync(x => x.Name == surveyConfiguration.Name);
            
            // dbContext.question_sets questionSets = new question_sets();

            var document = @"/home/microting/Documents/workspace/microting/ExcelReadManipPOC/Test-data.xlsx";
            using (FileStream fs = new FileStream(document, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    S sheets = doc.WorkbookPart.Workbook.Sheets;

                    foreach (E sheet in sheets)
                    {
                        foreach (A attr in sheet.GetAttributes())
                        {
                            Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
                        }
                    }

                    WorkbookPart workbookPart = doc.WorkbookPart;
                    WorksheetPart worksheetPart = GetWorksheetFromSheetName(workbookPart, "Senge_voksen");

                    var sheet1 = worksheetPart.Worksheet;
                    
                    
                    // Console.WriteLine(sheet1);
                

                    var cells = sheet1.Descendants<Cell>();
                    var rows = sheet1.Descendants<Row>();
                    var cols = sheet1.Descendants<Column>();

                    List<Column> columns = cols.ToList();

                    string text;
                    var rows1 = sheet1.GetFirstChild<SheetData>().Elements<Row>();
                    int i = 0;
                    var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>()
                        .FirstOrDefault();
                    List<KeyValuePair<string, sites>> localSites = new List<KeyValuePair<string, sites>>();
                    List<KeyValuePair<string, units>> localUnits = new List<KeyValuePair<string, units>>();
                    var languageId = dbContext.languages.SingleOrDefault(x => x.Name == "da-DK");
                    // List<sites> localSites = new List<sites>();
                    // List<units> localUnits = new List<units>();
                    foreach (var row in rows1)
                    {
                        if (i > 0)
                        {
                            var cells1 = row.Elements<Cell>();

                            int cellNumber = 0;
                            answers answer = null;
                            List<Cell> theCells = cells1.ToList();
                            var microtingUid = int.Parse(theCells[0].CellValue.Text);
                            var duration = int.Parse(theCells[1].CellValue.Text);
                            var date = stringTable.SharedStringTable.ElementAt(
                                int.Parse(theCells[2].CellValue.Text)).InnerText;
                            var time = stringTable.SharedStringTable.ElementAt(
                                int.Parse(theCells[7].CellValue.Text)).InnerText;
                            
                            DateTime dateOfDoing = DateTime.ParseExact($"{date} {time}","dd-MM-yyyy HH:mm:ss", null );
                            
                            var location = stringTable.SharedStringTable.ElementAt(
                                int.Parse(theCells[9].CellValue.Text)).InnerText;

                            int? sdkSiteId = null;
                            int? sdkUnitId = null;
                            
                            if (localSites.Any(x => x.Key == location))
                            {
                                sdkSiteId = localSites.First(x => x.Key == location).Value.Id;
                            }
                            else
                            {
                                var lookupSite = dbContext.sites.SingleOrDefault(x => x.Name == location);
                                if (lookupSite != null)
                                {
                                    KeyValuePair<string, sites> pair = new KeyValuePair<string, sites>(location, lookupSite);
                                    localSites.Add(pair);
                                    sdkSiteId = lookupSite.Id;
                                }
                                else
                                {
                                    sites site = new sites()
                                    {
                                        Name = location
                                    };
                                    await site.Create(dbContext);
                                    KeyValuePair<string, sites> pair = new KeyValuePair<string, sites>(location, site);
                                    localSites.Add(pair);
                                    sdkSiteId = site.Id;
                                }    
                            }
                            
                            var unitString = theCells[11].CellValue.Text;
                            if (localUnits.Any(x => x.Key == unitString))
                            {
                                sdkUnitId = localUnits.First(x => x.Key == unitString).Value.Id;
                            }
                            else
                            {
                                var lookupUnit = dbContext.units.SingleOrDefault(x => x.MicrotingUid.ToString() == unitString);
                            
                                if (lookupUnit != null)
                                {
                                    KeyValuePair<string, units> pair = new KeyValuePair<string, units>(unitString, lookupUnit);
                                    localUnits.Add(pair);
                                    sdkUnitId = lookupUnit.Id;
                                }
                                else
                                {
                                    units unit = new units()
                                    {
                                        MicrotingUid = int.Parse(unitString),
                                        SiteId = sdkSiteId
                                    };
                                    await unit.Create(dbContext);
                                    KeyValuePair<string, units> pair = new KeyValuePair<string, units>(unitString, unit);
                                    localUnits.Add(pair);
                                    sdkUnitId = unit.Id;
                                }
                            }



                            answer = dbContext.answers.SingleOrDefault(x =>
                                x.MicrotingUid == microtingUid);
                            if (answer == null)
                            {
                                answer = new answers()
                                {
                                    AnswerDuration = duration,
                                    UnitId = (int)sdkUnitId,
                                    SiteId = (int)sdkSiteId,
                                    MicrotingUid = microtingUid,
                                    FinishedAt = dateOfDoing,
                                    LanguageId = dbLanguage.Id,
                                    QuestionSetId = dbQuestionSets.Id,
                                    SurveyConfigurationId = dbSurveyConfiguration.Id
                                };   
                                await answer.Create(dbContext);
                            }
                            
                            foreach (var cell in cells1)
                            {
                                if (cell == null)
                                {
                                    Console.WriteLine("We got a null here");
                                }
                                else
                                {
                                    if (cellNumber > 12)
                                    {
                                        int questionLookupId = cellNumber;
                                        if (cellNumber > 25)
                                        {
                                            questionLookupId = 25;
                                        }

                                        int? lookupOptionId = null;
                                        if (cell.DataType != null)
                                        {
                                            //     if (cell.DataType.Value == CellValues.Number)
                                            //     {
                                            //         lookupOptionId = questionIds
                                            //             .First(x => x.Key == questionLookupId).Value.Options
                                            //             .SingleOrDefault(x => x.WeightValue == int.Parse(cell.CellValue.Text)).Id;
                                            if (cell.DataType.Value == CellValues.SharedString)
                                            {
                                                //         if (questionLookupId != 25)
                                                //         {
                                                foreach (options option in questionIds
                                                    .First(x => x.Key == questionLookupId).Value.Options)
                                                {
                                                    text = stringTable.SharedStringTable.ElementAt(
                                                        int.Parse(cell.CellValue.Text)).InnerText;

                                                    var r = option.OptionTranslationses.SingleOrDefault(x =>
                                                        x.Name == text);
                                                    if (r != null)
                                                    {
                                                        lookupOptionId = r.OptionId;
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (cellNumber > 13 && cellNumber < 25)
                                            {
                                                if (cell.CellValue != null)
                                                {
                                                    lookupOptionId = questionIds
                                                        .First(x => x.Key == questionLookupId).Value.Options
                                                        .SingleOrDefault(x => x.WeightValue == int.Parse(cell.CellValue.Text)).Id;    
                                                }
                                            }
                                            else
                                            {
                                                switch (cellNumber)
                                                {
                                                    case 25:
                                                        lookupOptionId = questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Options
                                                            .ToList()[0].Id;
                                                        break;
                                                    case 26:
                                                        lookupOptionId = questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Options
                                                            .ToList()[1].Id;
                                                        break;
                                                    case 27:
                                                        lookupOptionId = questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Options
                                                            .ToList()[2].Id;
                                                        break;
                                                    case 28:
                                                        lookupOptionId = questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Options
                                                            .ToList()[3].Id;
                                                        break;
                                                    case 29:
                                                        lookupOptionId = questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Options
                                                            .ToList()[4].Id;
                                                        break;
                                                }
                                            }
                                        }

                                        //         }
                                        //         else
                                        //         {
                                        //             if (questionLookupId == 25)
                                        //             {
                                        //                 lookupOptionId = questionIds.First(x => x.Key == questionLookupId)
                                        //                     .Value.Options.ToList()[cellNumber - 25].Id;
                                        //             }
                                        //         }
                                        //     }
                                        // }
                                        // else
                                        // {
                                        //     lookupOptionId = questionIds
                                        //         .First(x => x.Key == questionLookupId).Value.Options
                                        //         .SingleOrDefault(x => x.WeightValue == int.Parse(cell.CellValue.Text)).Id;
                                        // }
                                        //
                                        answer_values answerValue = null;
                                        if (lookupOptionId != null)
                                        {
                                            answerValue = dbContext.answer_values
                                                .SingleOrDefault(x 
                                                    => x.AnswerId == answer.Id 
                                                       && x.QuestionId == questionIds.First(y 
                                                           => y.Key == questionLookupId).Value.Id 
                                                       && x.OptionId == lookupOptionId);
                                        }
                                        else
                                        {
                                            answerValue = dbContext.answer_values
                                                .SingleOrDefault(x 
                                                    => x.AnswerId == answer.Id 
                                                       && x.QuestionId == questionIds.First(y 
                                                           => y.Key == questionLookupId).Value.Id);
                                        }
                                            
                                        if (answerValue == null)
                                        {
                                            answerValue = new answer_values()
                                            {
                                                AnswerId = answer.Id,
                                                QuestionId = questionIds.First(x => x.Key == questionLookupId).Value.Id
                                            };

                                            if (cell.DataType != null)
                                            {
                                                if (cell.DataType.Value == CellValues.SharedString)
                                                {
                                                    if (stringTable != null)
                                                    {
                                                        text = stringTable.SharedStringTable.ElementAt(
                                                            int.Parse(cell.CellValue.Text)).InnerText;
                                                        // Console.WriteLine(text + " ");
                                                        int optionId = 0;
                                                        foreach (options option in questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Options)
                                                        {
                                                            var r =  option.OptionTranslationses.SingleOrDefault(x => x.Name == text);
                                                            if (r != null)
                                                            {
                                                                optionId = r.OptionId;
                                                            }
                                                        }
                                                        answerValue.Value = text;
                                                        answerValue.OptionId = optionId;
                                                        answerValue.QuestionId = questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Id;
                                                        await answerValue.Create(dbContext);
                                                    }    
                                                }
                                                
                                            }
                                            else
                                            {
                                                if (cellNumber > 13 && cellNumber < 25)
                                                {
                                                    if (cell.CellValue != null)
                                                    {
                                                        text = cell.CellValue.Text;
                                                    
                                                        answerValue.Value = text;
                                                        answerValue.QuestionId = questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Id;
                                                        answerValue.OptionId = questionIds
                                                            .First(x => x.Key == questionLookupId).Value.Options
                                                            .SingleOrDefault(x => x.WeightValue == int.Parse(text)).Id;
                                                        await answerValue.Create(dbContext);
                                                        // Console.WriteLine(cell.CellValue.Text);
                                                    }
                                                }
                                                else
                                                {
                                                    if (cell.CellValue != null)
                                                    {
                                                        if (int.Parse(cell.CellValue.Text) == 1)
                                                        {
                                                            answerValue.Value = "1";
                                                            answerValue.QuestionId = questionIds
                                                                .First(x => x.Key == questionLookupId).Value.Id;
                                                            switch (cellNumber)
                                                            {
                                                                case 25:
                                                                    answerValue.OptionId = questionIds
                                                                        .First(x => x.Key == questionLookupId).Value.Options.ToList()[0].Id;
                                                                    break;
                                                                case 26:
                                                                    answerValue.OptionId = questionIds
                                                                        .First(x => x.Key == questionLookupId).Value.Options.ToList()[1].Id;
                                                                    break;
                                                                case 27:
                                                                    answerValue.OptionId = questionIds
                                                                        .First(x => x.Key == questionLookupId).Value.Options.ToList()[2].Id;
                                                                    break;
                                                                case 28:
                                                                    answerValue.OptionId = questionIds
                                                                        .First(x => x.Key == questionLookupId).Value.Options.ToList()[3].Id;
                                                                    break;
                                                                case 29:
                                                                    answerValue.OptionId = questionIds
                                                                        .First(x => x.Key == questionLookupId).Value.Options.ToList()[4].Id;
                                                                    break;
                                                            }
                                                            
                                                            await answerValue.Create(dbContext);
                                                        }   
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                cellNumber++;
                            }
                        }
                        Console.WriteLine($"ROW {i}");

                        i++;

                    }
                    
                    
                    
                    // Console.WriteLine("Row count = {0}", rows.First());
                    // Console.WriteLine("Cell count = {0}", cells.LongCount());
                    // Console.WriteLine("Column count = {0}", columns.LongCount());
                }
                
            }
        }
        private static WorksheetPart GetWorksheetFromSheetName(WorkbookPart workbookPart, string sheetName)
        {     
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null)
            {
                throw new Exception(string.Format("Could not find sheet with name {0}", sheetName));
            }
            else
            {
                return workbookPart.GetPartById(sheet.Id) as WorksheetPart;
            }
        }
        
        
    }
}