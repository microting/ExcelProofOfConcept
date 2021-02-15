using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using eFormCore;
using Microsoft.EntityFrameworkCore;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using Microting.eForm.Infrastructure;
using Microting.eForm.Infrastructure.Data.Entities;
using Constants = Microting.eForm.Infrastructure.Constants.Constants;

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
                    "host= localhost;Database=420_SDK;user = root; password=secretpassword;port=3306;Convert Zero Datetime = true;SslMode=none;")
                .Result;

            MicrotingDbContext dbContext = _core.DbContextHelper.GetDbContext();

            QuestionSet questionSets = new QuestionSet
            {
                Name = "Test-Set"
            };
            if (dbContext.QuestionSets.Count(x => x.Name == questionSets.Name) != 1)
            {
               await questionSets.Create(dbContext);
            }

            Language language = new Language
            {
                LanguageCode = "Description",
                Name = "Danish"
            };
            if (dbContext.Languages.Count(x => x.Name == "da-DK") != 1)
            {
                await language.Create(dbContext);
            }

            Language dbLanguage = await dbContext.Languages.FirstOrDefaultAsync(x => x.Name == language.Name);

            QuestionSet dbQuestionSets = await dbContext.QuestionSets.FirstOrDefaultAsync(x => x.Name == questionSets.Name);

            string[] questionNames = new[] {"Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8", "Q9", "Q10", "Q11", "Q12", "Q13"};
            //                                13    14    15    16    17    18    19    20    21    22    23    24
            List<KeyValuePair<int, Question>> questionIds = new List<KeyValuePair<int, Question>>();

            int qi = 13;
            foreach (var questionName in questionNames)
            {
                if (questionName != "Q13" && questionName != "Q1")
                {
                    var questionTranslation =
                        dbContext.QuestionTranslations.SingleOrDefault(x => x.Name == questionName);

                    if (questionTranslation == null)
                    {
                        Question question = new Question()
                        {
                            QuestionSetId = dbQuestionSets.Id,
                            QuestionType = Constants.QuestionTypes.Smiley2
                        };
                        await question.Create(dbContext);

                        KeyValuePair<int, Question> kvp = new KeyValuePair<int, Question>(qi, question);
                        questionIds.Add(kvp);

                        questionTranslation = new QuestionTranslation()
                        {
                            Name = questionName,
                            QuestionId = question.Id,
                            LanguageId = dbLanguage.Id
                        };
                        await questionTranslation.Create(dbContext);
                    }
                    else
                    {
                        KeyValuePair<int, Question> kvp = new KeyValuePair<int, Question>(qi, questionTranslation.Question);
                        questionIds.Add(kvp);
                    }
                }
                else
                {
                    var questionTranslation =
                        dbContext.QuestionTranslations.SingleOrDefault(x => x.Name == questionName);

                    if (questionTranslation == null)
                    {
                        Question question = new Question()
                        {
                            QuestionSetId = dbQuestionSets.Id,
                            QuestionType = questionName == "Q1" ? Constants.QuestionTypes.List : Constants.QuestionTypes.Multi
                        };
                        await question.Create(dbContext);

                        questionTranslation = new QuestionTranslation()
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
                            Option option = new Option()
                            {
                                QuestionId = question.Id,
                                Weight = 1,
                                WeightValue = 1
                            };
                            await option.Create(dbContext);

                            OptionTranslation optionTranslation = new OptionTranslation()
                            {
                                OptionId = option.Id,
                                Name = questionOption,
                                LanguageId = dbLanguage.Id
                            };

                            await optionTranslation.Create(dbContext);
                        }

                        KeyValuePair<int, Question> kvp = new KeyValuePair<int, Question>(qi, question);
                        questionIds.Add(kvp);
                    }
                    else
                    {
                        KeyValuePair<int, Question> kvp = new KeyValuePair<int, Question>(qi, questionTranslation.Question);
                        questionIds.Add(kvp);
                    }
                }
                qi++;
            }

            // Q13 with options

            // KeyValuePair<int, questions> kvp = new KeyValuePair<int, questions>(qi, questionTranslation.Question);
            // questionIds.Add(kvp);

            SurveyConfiguration surveyConfiguration = new SurveyConfiguration
            {
                QuestionSetId = dbQuestionSets.Id,
                Name = "Configuartion 1"
            };
            if (dbContext.SurveyConfigurations.Count(x => x.Name == surveyConfiguration.Name) != 1)
            {
                await surveyConfiguration.Create(dbContext);
            }

            SurveyConfiguration dbSurveyConfiguration =
                await dbContext.SurveyConfigurations.FirstOrDefaultAsync(x => x.Name == surveyConfiguration.Name);

            // dbContext.question_sets questionSets = new question_sets();
            Random rnd = new Random();
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
                    List<KeyValuePair<string, Site>> localSites = new List<KeyValuePair<string, Site>>();
                    List<KeyValuePair<string, Unit>> localUnits = new List<KeyValuePair<string, Unit>>();
                    var languageId = dbContext.Languages.SingleOrDefault(x => x.Name == "da-DK");
                    // List<sites> localSites = new List<sites>();
                    // List<units> localUnits = new List<units>();
                    foreach (var row in rows1)
                    {
                        if (i > 0 && i < 553)
                        {
                            var cells1 = row.Elements<Cell>();

                            int cellNumber = 0;
                            Answer answer = null;
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
                                var lookupSite = dbContext.Sites.SingleOrDefault(x => x.Name == location);
                                if (lookupSite != null)
                                {
                                    KeyValuePair<string, Site> pair = new KeyValuePair<string, Site>(location, lookupSite);
                                    localSites.Add(pair);
                                    sdkSiteId = lookupSite.Id;
                                }
                                else
                                {
                                    Site site = new Site
                                    {
                                        Name = location,
                                        MicrotingUid = rnd.Next(1, 999999)
                                    };
                                    await site.Create(dbContext);
                                    KeyValuePair<string, Site> pair = new KeyValuePair<string, Site>(location, site);
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
                                var lookupUnit = dbContext.Units.SingleOrDefault(x => x.MicrotingUid.ToString() == unitString);

                                if (lookupUnit != null)
                                {
                                    KeyValuePair<string, Unit> pair = new KeyValuePair<string, Unit>(unitString, lookupUnit);
                                    localUnits.Add(pair);
                                    sdkUnitId = lookupUnit.Id;
                                }
                                else
                                {
                                    Unit unit = new Unit
                                    {
                                        MicrotingUid = int.Parse(unitString),
                                        SiteId = sdkSiteId
                                    };
                                    await unit.Create(dbContext);
                                    KeyValuePair<string, Unit> pair = new KeyValuePair<string, Unit>(unitString, unit);
                                    localUnits.Add(pair);
                                    sdkUnitId = unit.Id;
                                }
                            }

                            answer = dbContext.Answers.SingleOrDefault(x =>
                                x.MicrotingUid == microtingUid);
                            if (answer == null)
                            {
                                answer = new Answer
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
                                                foreach (Option option in questionIds
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
                                        AnswerValue answerValue = null;
                                        if (lookupOptionId != null)
                                        {
                                            answerValue = dbContext.AnswerValues
                                                .SingleOrDefault(x
                                                    => x.AnswerId == answer.Id
                                                       && x.QuestionId == questionIds.First(y
                                                           => y.Key == questionLookupId).Value.Id
                                                       && x.OptionId == lookupOptionId);
                                        }
                                        else
                                        {
                                            answerValue = dbContext.AnswerValues
                                                .SingleOrDefault(x
                                                    => x.AnswerId == answer.Id
                                                       && x.QuestionId == questionIds.First(y
                                                           => y.Key == questionLookupId).Value.Id);
                                        }

                                        if (answerValue == null)
                                        {
                                            answerValue = new AnswerValue
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
                                                        foreach (Option option in questionIds
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
                }
            }
            // Q1
            int q1Id = 1;
            int q2Id = 2;
            int q3Id = 3;
            int q4Id = 4;
            int q5Id = 5;
            int q6Id = 6;
            int q7Id = 7;
            int q8Id = 8;
            int q9Id = 9;
            int q10Id = 10;
            int q11Id = 11;
            int q12Id = 12;
            int q13Id = 13;
            int optionJaId = 1;
            int optionNejId = 2;
            int option100q2Id = 3;
            int option75q2Id = 4;
            int option50q2Id = 5;
            int option25q2Id = 6;
            int option0q2Id = 7;
            int option999q2Id = 8;
            int option100q3Id = 9;
            int option75q3Id = 10;
            int option50q3Id = 11;
            int option25q3Id = 12;
            int option0q3Id = 13;
            int option999q3Id = 14;
            int option100q4Id = 15;
            int option75q4Id = 16;
            int option50q4Id = 17;
            int option25q4Id = 18;
            int option0q4Id = 19;
            int option999q4Id = 20;
            int option100q5Id = 21;
            int option75q5Id = 22;
            int option50q5Id = 23;
            int option25q5Id = 24;
            int option0q5Id = 25;
            int option999q5Id = 26;
            int option100q6Id = 27;
            int option75q6Id = 28;
            int option50q6Id = 29;
            int option25q6Id = 30;
            int option0q6Id = 31;
            int option999q6Id = 32;
            int option100q7Id = 33;
            int option75q7Id = 34;
            int option50q7Id = 35;
            int option25q7Id = 36;
            int option0q7Id = 37;
            int option999q7Id = 38;
            int option100q8Id = 39;
            int option75q8Id = 40;
            int option50q8Id = 41;
            int option25q8Id = 42;
            int option0q8Id = 43;
            int option999q8Id = 44;
            int option100q9Id = 45;
            int option75q9Id = 46;
            int option50q9Id = 47;
            int option25q9Id = 48;
            int option0q9Id = 49;
            int option999q9Id = 50;
            int option100q10Id = 51;
            int option75q10Id = 52;
            int option50q10Id = 53;
            int option25q10Id = 54;
            int option0q10Id = 55;
            int option999q10Id = 56;
            int option100q11Id = 57;
            int option75q11Id = 58;
            int option50q11Id = 59;
            int option25q11Id = 60;
            int option0q11Id = 61;
            int option999q11Id = 62;
            int option100q12Id = 63;
            int option75q12Id = 64;
            int option50q12Id = 65;
            int option25q12Id = 66;
            int option0q12Id = 67;
            int option999q12Id = 68;
            int optionq13_1Id = 69;
            int optionq13_2Id = 70;
            int optionq13_3Id = 71;
            int optionq13_4Id = 72;
            int optionq13_5Id = 73;
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == optionJaId) == 419);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == optionNejId) == 133);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q1Id) == 552);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q2Id) == 5);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q2Id) == 7);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q2Id) == 14);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q2Id) == 112);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q2Id) == 275);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q2Id) == 6);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q2Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q3Id) == 15);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q3Id) == 8);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q3Id) == 44);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q3Id) == 144);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q3Id) == 201);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q3Id) == 7);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q3Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q4Id) == 13);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q4Id) == 17);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q4Id) == 78);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q4Id) == 123);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q4Id) == 176);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q4Id) == 12);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q4Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q5Id) == 16);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q5Id) == 18);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q5Id) == 49);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q5Id) == 135);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q5Id) == 188);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q5Id) == 13);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q5Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q6Id) == 21);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q6Id) == 23);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q6Id) == 61);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q6Id) == 131);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q6Id) == 160);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q6Id) == 23);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q6Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q7Id) == 13);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q7Id) == 8);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q7Id) == 57);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q7Id) == 116);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q7Id) == 216);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q7Id) == 9);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q7Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q8Id) == 35);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q8Id) == 27);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q8Id) == 98);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q8Id) == 108);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q8Id) == 124);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q8Id) == 27);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q8Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q9Id) == 19);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q9Id) == 23);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q9Id) == 51);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q9Id) == 107);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q9Id) == 213);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q9Id) == 6);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q9Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q10Id) == 16);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q10Id) == 10);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q10Id) == 66);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q10Id) == 116);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q10Id) == 186);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q10Id) == 25);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q10Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q11Id) == 11);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q11Id) == 8);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q11Id) == 41);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q11Id) == 111);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q11Id) == 211);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q11Id) == 37);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q11Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option0q12Id) == 12);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option25q12Id) == 9);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option50q12Id) == 58);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option75q12Id) == 126);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option100q12Id) == 187);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == option999q12Id) == 27);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q12Id) == 419);

            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == optionq13_1Id) == 289);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q13Id) == 1383);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == optionq13_2Id) == 273);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q13Id) == 1383);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == optionq13_3Id) == 281);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q13Id) == 1383);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == optionq13_4Id) == 271);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q13Id) == 1383);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.OptionId == optionq13_5Id) == 269);
            Debug.Assert(dbContext.AnswerValues.Count(x => x.QuestionId == q13Id) == 1383);
            Console.WriteLine("we are done");
        }
        private static WorksheetPart GetWorksheetFromSheetName(WorkbookPart workbookPart, string sheetName)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null)
            {
                throw new Exception(string.Format("Could not find sheet with name {0}", sheetName));
            }

            return workbookPart.GetPartById(sheet.Id) as WorksheetPart;
        }
    }
}