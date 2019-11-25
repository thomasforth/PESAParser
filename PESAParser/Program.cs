using CsvHelper;
using CsvHelper.Configuration.Attributes;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace PESAParser
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            List<PESA2005Entry> PESA2005Entries = ParsePESA2005();
            List<PESA2010Chap9Entry> PESA2010EntriesChap9 = ParsePESADBs2010Chap9();
            List<PESA2010Entry> PESA2010EntriesChap10 = ParsePESADBs2010Chap10();
            List<PESA2014Entry> PESA2014Entries = ParsePESADB2014();
            List<PESA2018Entry> PESA2018Entries = ParsePESADB2018();

            // Create geographies standardising dictionaries
            List<StandardGeography> standardGeographies = new List<StandardGeography>();
            using (TextReader textReader = File.OpenText("Assets/StandardGeographyNames.csv"))
            {
                CsvReader csvReader = new CsvReader(textReader);
                standardGeographies = csvReader.GetRecords<StandardGeography>().ToList();
            }

            Dictionary<string, string> standardGeographiesDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> standardGeographiesCodeDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach(StandardGeography standardGeography in standardGeographies)
            {
                standardGeographiesDict.Add(standardGeography.Name, standardGeography.StandardName);
                standardGeographiesCodeDict.Add(standardGeography.Name, standardGeography.GeoCode);
            }

            // Create function standardising dictionary
            List<StandardFunctionName> standardFunctionNames = new List<StandardFunctionName>();
            using (TextReader textReader = File.OpenText("Assets/StandardFunctionNames.csv"))
            {
                CsvReader csvReader = new CsvReader(textReader);
                standardFunctionNames = csvReader.GetRecords<StandardFunctionName>().ToList();
            }
            Dictionary<string, string> standardFunctionsDictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach(StandardFunctionName standardFunctionName in standardFunctionNames)
            {
                standardFunctionsDictionary.Add(standardFunctionName.Function, standardFunctionName.StandardFunction);
            }

            // Create subfunction standardising dictionary
            List<StandardSubfunctionName> standardSubfunctionNames = new List<StandardSubfunctionName>();
            using (TextReader textReader = File.OpenText("Assets/StandardSubfunctionNames.csv"))
            {
                CsvReader csvReader = new CsvReader(textReader);
                standardSubfunctionNames = csvReader.GetRecords<StandardSubfunctionName>().ToList();
            }
            Dictionary<string, string> standardSubfunctionsDictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (StandardSubfunctionName standardSubfunctionName in standardSubfunctionNames)
            {
                standardSubfunctionsDictionary.Add(standardSubfunctionName.Subfunction, standardSubfunctionName.StandardSubfunction);
            }

            // Create deflator dictionary
            List<Inflation> Inflators = new List<Inflation>();
            using (TextReader textReader = File.OpenText("Assets/CPI_inflation_Nov_2019.csv"))
            {
                CsvReader csvReader = new CsvReader(textReader);
                Inflators = csvReader.GetRecords<Inflation>().ToList();
            }
            Dictionary<int, double> InflatorDictionary = new Dictionary<int, double>();
            foreach(Inflation inflation in Inflators)
            {
                InflatorDictionary.Add(inflation.Year, inflation.CPI2015);
            }

            // Create population dictionary
            List<Population> Populations = new List<Population>();
            using (TextReader textReader = File.OpenText("Assets/demo_r_d2jan_1_Data.csv"))
            {
                CsvReader csvReader = new CsvReader(textReader);
                csvReader.Configuration.HeaderValidated = null;
                csvReader.Configuration.MissingFieldFound = null;
                Populations = csvReader.GetRecords<Population>().ToList();
            }            

            List<CombinedPESAEntry> CombinedPESAEntries = new List<CombinedPESAEntry>();
            foreach(PESA2005Entry entry in PESA2005Entries)
            {
                CombinedPESAEntry combinedPESAEntry = new CombinedPESAEntry()
                {                    
                    CAPorCUR = entry.CAPorCUR,
                    FinancialYear = entry.FinancialYear,
                    Geography = standardGeographiesDict[entry.Geography],
                    HMTFunction = standardFunctionsDictionary[entry.Function.Replace("\n", "")],
                    Value = 1000000 * entry.Value,
                    YearEnd = entry.YearEnd,
                    CPI2015 = InflatorDictionary[entry.YearEnd],
                };

                if (entry.YearEnd < 2005 && entry.Function != "Total")
                {
                    CombinedPESAEntries.Add(combinedPESAEntry);
                }
            }
            foreach(PESA2010Chap9Entry entry in PESA2010EntriesChap9)
            {
                CombinedPESAEntry combinedPESAEntry = new CombinedPESAEntry()
                {
                    CAPorCUR = entry.CAPorCUR,
                    FinancialYear = entry.FinancialYear,
                    Geography = standardGeographiesDict[entry.NUTSRegion],
                    HMTFunction = standardFunctionsDictionary[entry.HMTFunction.Replace("\n", "")],
                    Value = 1000000 * entry.Value,
                    YearEnd = entry.YearEnd,
                    CPI2015 = InflatorDictionary[entry.YearEnd],
                };
                CombinedPESAEntries.Add(combinedPESAEntry);
            }
            foreach (PESA2010Entry entry in PESA2010EntriesChap10)
            {
                CombinedPESAEntry combinedPESAEntry = new CombinedPESAEntry()
                {
                    CAPorCUR = entry.CAPorCUR,
                    FinancialYear = entry.FinancialYear,
                    Geography = standardGeographiesDict[entry.NUTSRegion],
                    HMTFunction = standardFunctionsDictionary[entry.HMTFunction.Replace("\n", "")],
                    Value = 1000000 * entry.Value,
                    YearEnd = entry.YearEnd,
                    CPI2015 = InflatorDictionary[entry.YearEnd],
                    HMTSubfunction = entry.HMTSubfunction.Replace("\n", "")
                };
                CombinedPESAEntries.Add(combinedPESAEntry);
            }
            foreach (PESA2014Entry entry in PESA2014Entries)
            {
                CombinedPESAEntry combinedPESAEntry = new CombinedPESAEntry()
                {
                    CAPorCUR = entry.CAPorCUR,
                    FinancialYear = entry.FinancialYear,
                    Geography = standardGeographiesDict[entry.NUTSRegion],
                    HMTFunction = standardFunctionsDictionary[entry.HMTFunction.Replace("\n", "")],
                    Value = 1000 * entry.Value,
                    YearEnd = entry.YearEnd,
                    CPI2015 = InflatorDictionary[entry.YearEnd],
                    HMTSubfunction = entry.HMTSubfunction.Replace("\n", "")
                };
                CombinedPESAEntries.Add(combinedPESAEntry);
            }
            foreach (PESA2018Entry entry in PESA2018Entries)
            {
                CombinedPESAEntry combinedPESAEntry = new CombinedPESAEntry()
                {
                    CAPorCUR = entry.CAPorCUR,
                    FinancialYear = entry.FinancialYear,
                    Geography = standardGeographiesDict[entry.NUTSRegion],
                    HMTFunction = standardFunctionsDictionary[entry.HMTFunction.Replace("\n", "")],
                    Value = 1000 * entry.Value,
                    YearEnd = entry.YearEnd,
                    CPI2015 = InflatorDictionary[entry.YearEnd],
                    HMTSubfunction = entry.HMTSubfunction.Replace("\n", "")
                };

                CombinedPESAEntries.Add(combinedPESAEntry);
            }

            foreach(CombinedPESAEntry combinedPESAEntry in CombinedPESAEntries)
            {
                combinedPESAEntry.GeographyCode = standardGeographiesCodeDict[combinedPESAEntry.Geography];
                if (Populations.Where(x => x.GEO == combinedPESAEntry.GeographyCode && x.TIME == combinedPESAEntry.YearEnd).FirstOrDefault() != null)
                {
                    combinedPESAEntry.Population = Populations.Where(x => x.GEO == combinedPESAEntry.GeographyCode && x.TIME == combinedPESAEntry.YearEnd).FirstOrDefault().Value;
                }
                combinedPESAEntry.Value2015PerCapita = (combinedPESAEntry.Value * (100 / combinedPESAEntry.CPI2015)) / combinedPESAEntry.Population;
                if (combinedPESAEntry.HMTSubfunction != null)
                {
                    combinedPESAEntry.HMTSubfunction = standardSubfunctionsDictionary[combinedPESAEntry.HMTSubfunction];
                }
            }

            using (TextWriter TextWriter = File.CreateText(@"ParsedCombinedPESA.csv"))
            {
                CsvWriter CSVwriter = new CsvWriter(TextWriter);
                CSVwriter.WriteRecords(CombinedPESAEntries);
            }
        }

        static List<PESA2005Entry> ParsePESA2005()
        {
            string ExcelFilePath = @"Assets/pesa2005_chapter8_tablesv3.xls";

            //List<string> RegionNames = new List<string>() { "North East", "North West", "Yorkshire and Humberside", "East Midlands", "West Midlands", "South West", "Eastern", "London", "South East", "Total England", "Scotland", "Wales", "Northern Ireland", "UK Identifiable expenditure", "Outside UK", "Total Identifiable expenditure" };
            List<string> ListOfSheetsToParse = new List<string>() { "8.5a", "8.5b", "8.6a", "8.6b", "8.7a", "8.7b", "8.8a", "8.8b", "8.9a", "8.9b", "8.10a", "8.10b" };

            List<PESA2005Entry> PESAEntries = new List<PESA2005Entry>();
            foreach (string SheetName in ListOfSheetsToParse)
            {
                DataTable CURRENTTABLE;
                using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        DataTableCollection Worksheets = result.Tables;
                        CURRENTTABLE = Worksheets[Worksheets.IndexOf(SheetName)];
                    }
                }

                string Title = (string)CURRENTTABLE.Rows[0].ItemArray[0];

                List<string> Units = CURRENTTABLE.Rows[2].ItemArray.OfType<string>().ToList();
                string Unit = Units.Last();

                List<string> Headers = CURRENTTABLE.Rows[3].ItemArray.OfType<string>().ToList();

                List<PESA2005Entry> PESAEntriesForThisSheet = new List<PESA2005Entry>();
                for (int i = 4; i < 22; i++)
                {
                    foreach (string function in Headers)
                    {
                        if (CURRENTTABLE.Rows[i].ItemArray[0] != System.DBNull.Value)
                        {                            
                            PESA2005Entry pesaEntry = new PESA2005Entry();
                            pesaEntry.Geography = (string)CURRENTTABLE.Rows[i].ItemArray[0];
                            pesaEntry.Function = function.Trim();
                            pesaEntry.TableTitle = Title;
                            pesaEntry.Unit = Unit;
                            pesaEntry.FinancialYear = Title.Split(",").Last().Trim();
                            pesaEntry.YearEnd = int.Parse(string.Join("", pesaEntry.FinancialYear.Take(4))) + 1;
                            pesaEntry.Value = SafeConvertObjectToInt(CURRENTTABLE.Rows[i].ItemArray[Headers.IndexOf(function) + 1]);

                            if (Title.Contains("current"))
                            {
                                pesaEntry.CAPorCUR = "CUR";
                            }
                            if (Title.Contains("capital"))
                            {
                                pesaEntry.CAPorCUR = "CAP";
                            }

                            PESAEntriesForThisSheet.Add(pesaEntry);
                        }
                    }
                }
                PESAEntries.AddRange(PESAEntriesForThisSheet);
            }
            return PESAEntries;
        }


        static List<PESA2010Chap9Entry> ParsePESADBs2010Chap9()
        {
            string ExcelFilePath = @"Assets/pesa_2010_database_tables_chapter9.xlsx";
            string SheetName = @"CRA 2010 Chapter 9 DB final ";

            List<PESA2010Chap9Entry> PESA2010Chap9Entries = new List<PESA2010Chap9Entry>();

            DataTable CURRENTTABLE;
            using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTableCollection Worksheets = result.Tables;
                    CURRENTTABLE = Worksheets[Worksheets.IndexOf(SheetName)];
                }
            }

            List<string> FinancialYears = new List<string>()
            {
                "2004-05",
                "2005-06",
                "2006-07",
                "2007-08",
                "2008-09",
                "2009-10"
            };
            FinancialYears.Remove("2009-10"); // already covered in a later dataset

            for (int i = 1; i < CURRENTTABLE.Rows.Count; i++)
            {
                foreach (string FinancialYear in FinancialYears)
                {
                    PESA2010Chap9Entry pesaEntry = new PESA2010Chap9Entry();

                    pesaEntry.FinancialYear = FinancialYear;
                    pesaEntry.YearEnd = int.Parse(FinancialYear.Substring(0, 4)) + 1;
                    pesaEntry.Value = SafeConvertObjectToDouble(CURRENTTABLE.Rows[i].ItemArray[CURRENTTABLE.Rows[0].ItemArray.OfType<string>().ToList().IndexOf(FinancialYear)]);

                    pesaEntry.DepartmentCode = SafeConvertObjectToString(CURRENTTABLE.Rows[i].ItemArray[0]);
                    pesaEntry.DepartmentName = (string)CURRENTTABLE.Rows[i].ItemArray[1];
                    pesaEntry.COFOGLevel1 = (string)CURRENTTABLE.Rows[i].ItemArray[2];
                    pesaEntry.HMTFunction = (string)CURRENTTABLE.Rows[i].ItemArray[3];
                    pesaEntry.ProgrammeObjectGroup = SafeConvertObjectToString(CURRENTTABLE.Rows[i].ItemArray[4]);
                    pesaEntry.ProgrammeObjectGroupAlias = SafeConvertObjectToString(CURRENTTABLE.Rows[i].ItemArray[5]);
                    pesaEntry.IDNonID = (string)CURRENTTABLE.Rows[i].ItemArray[6];
                    pesaEntry.CAPorCUR = (string)CURRENTTABLE.Rows[i].ItemArray[7];
                    pesaEntry.CGorLGorPCorBOE = (string)CURRENTTABLE.Rows[i].ItemArray[8];
                    pesaEntry.NUTSRegion = (string)CURRENTTABLE.Rows[i].ItemArray[9];

                    PESA2010Chap9Entries.Add(pesaEntry);
                }
            }
            return PESA2010Chap9Entries;
        }


        static List<PESA2010Entry> ParsePESADBs2010Chap10()
        {
            string ExcelFilePath = @"Assets/pesa_2010_database_tables_chapter10.xlsx";
            string SheetName = @"CRA 2010 Chapter 10 DB final ";

            List<PESA2010Entry> PESA2010Entries = new List<PESA2010Entry>();

            DataTable CURRENTTABLE;
            using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTableCollection Worksheets = result.Tables;
                    CURRENTTABLE = Worksheets[Worksheets.IndexOf(SheetName)];
                }
            }

            List<string> FinancialYears = new List<string>()
            {
                "2004-05",
                "2005-06",
                "2006-07", 
                "2007-08",
                "2008-09",
                "2009-10"
            };
            FinancialYears.Remove("2009-10"); // already covered in a later dataset

            for (int i = 1; i < CURRENTTABLE.Rows.Count; i++)
            {
                foreach (string FinancialYear in FinancialYears)
                {
                    PESA2010Entry pesaEntry = new PESA2010Entry();

                    pesaEntry.FinancialYear = FinancialYear;
                    pesaEntry.YearEnd = int.Parse(FinancialYear.Substring(0, 4)) + 1;
                    pesaEntry.Value = SafeConvertObjectToDouble(CURRENTTABLE.Rows[i].ItemArray[CURRENTTABLE.Rows[0].ItemArray.OfType<string>().ToList().IndexOf(FinancialYear)]);

                    pesaEntry.DepartmentCode = SafeConvertObjectToString(CURRENTTABLE.Rows[i].ItemArray[0]);
                    pesaEntry.DepartmentName = (string)CURRENTTABLE.Rows[i].ItemArray[1];
                    pesaEntry.COFOGLevel1 = (string)CURRENTTABLE.Rows[i].ItemArray[2];
                    pesaEntry.HMTFunction = (string)CURRENTTABLE.Rows[i].ItemArray[3];
                    pesaEntry.COFOGLevel2 = (string)CURRENTTABLE.Rows[i].ItemArray[4];
                    pesaEntry.HMTSubfunction = (string)CURRENTTABLE.Rows[i].ItemArray[5];
                    pesaEntry.ProgrammeObjectGroup = SafeConvertObjectToString(CURRENTTABLE.Rows[i].ItemArray[6]);
                    pesaEntry.ProgrammeObjectGroupAlias = SafeConvertObjectToString(CURRENTTABLE.Rows[i].ItemArray[7]);
                    pesaEntry.IDNonID = (string)CURRENTTABLE.Rows[i].ItemArray[8];
                    pesaEntry.CAPorCUR = (string)CURRENTTABLE.Rows[i].ItemArray[9];
                    pesaEntry.NUTSRegion = (string)CURRENTTABLE.Rows[i].ItemArray[11];
                    pesaEntry.CGorLGorPCorBOE = (string)CURRENTTABLE.Rows[i].ItemArray[10];

                    PESA2010Entries.Add(pesaEntry);
                }
            }
            return PESA2010Entries;
        }

        static List<PESA2014Entry> ParsePESADB2014()
        {
            string ExcelFilePath = @"Assets/CRA_2014_Combined_Database_for_Publication.xlsx";
            string SheetName = @"CRA14 combined database";

            DataTable CURRENTTABLE;
            using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTableCollection Worksheets = result.Tables;
                    CURRENTTABLE = Worksheets[Worksheets.IndexOf(SheetName)];
                }
            }

            List<PESA2014Entry> PESA2014Entries = new List<PESA2014Entry>();

            List<string> FinancialYears = new List<string>()
            {
                "2009-10",
                "2010-11",
                "2011-12",
                "2012-13",
                "2013-14"
            };
            FinancialYears.Remove("2013-14"); // already covered in a later dataset

            for (int i = 1; i < CURRENTTABLE.Rows.Count; i++)
            {
                foreach (string FinancialYear in FinancialYears)
                {
                    PESA2014Entry pesaEntry = new PESA2014Entry();

                    pesaEntry.FinancialYear = FinancialYear;
                    pesaEntry.YearEnd = int.Parse(FinancialYear.Substring(0, 4)) + 1;
                    pesaEntry.Value = SafeConvertObjectToDouble(CURRENTTABLE.Rows[i].ItemArray[CURRENTTABLE.Rows[0].ItemArray.OfType<string>().ToList().IndexOf(FinancialYear)]);

                    // Department COFOG Level 0   HMT Function    COFOG Level 1   HMT Subfunction CRA Segment CRA Segment Description ID/ non - ID   CAP or CUR CG, LG or PC    NUTS Region Country

                    pesaEntry.DepartmentName = (string)CURRENTTABLE.Rows[i].ItemArray[0];
                    //pesaEntry.OrganisationName = (string)CURRENTTABLE.Rows[i].ItemArray[1];
                    pesaEntry.CRASegmentCode = (string)CURRENTTABLE.Rows[i].ItemArray[5];
                    pesaEntry.CRASegmentName = (string)CURRENTTABLE.Rows[i].ItemArray[6];
                    pesaEntry.COFOGLevel0 = (string)CURRENTTABLE.Rows[i].ItemArray[1];
                    pesaEntry.HMTFunction = (string)CURRENTTABLE.Rows[i].ItemArray[2];
                    pesaEntry.COFOGLevel1 = (string)CURRENTTABLE.Rows[i].ItemArray[3];
                    pesaEntry.HMTSubfunction = (string)CURRENTTABLE.Rows[i].ItemArray[4];
                    pesaEntry.IDNonID = (string)CURRENTTABLE.Rows[i].ItemArray[7];
                    pesaEntry.CAPorCUR = (string)CURRENTTABLE.Rows[i].ItemArray[8];
                    pesaEntry.CGorLGorPCorBOE = (string)CURRENTTABLE.Rows[i].ItemArray[9];
                    //pesaEntry.HMTorDept = (string)CURRENTTABLE.Rows[i].ItemArray[11];
                    pesaEntry.NUTSRegion = (string)CURRENTTABLE.Rows[i].ItemArray[10];
                    pesaEntry.Country = (string)CURRENTTABLE.Rows[i].ItemArray[11];

                    PESA2014Entries.Add(pesaEntry);
                }
            }
            return PESA2014Entries;
        }

        static List<PESA2018Entry> ParsePESADB2018()
        {
            string ExcelFilePath = @"Assets/CRA_2018_Database_for_Publication_rvsd.xlsx";
            string SheetName = @"CRA 2018 database";

            DataTable CURRENTTABLE;
            using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTableCollection Worksheets = result.Tables;
                    CURRENTTABLE = Worksheets[Worksheets.IndexOf(SheetName)];
                }
            }

            List<PESA2018Entry> PESA2018Entries = new List<PESA2018Entry>();

            List<string> FinancialYears = new List<string>()
            {
                "2013-14",
                "2014-15",
                "2015-16",
                "2016-17",
                "2017-18"
            };

            for (int i = 1; i < CURRENTTABLE.Rows.Count; i++)
            {
                foreach (string FinancialYear in FinancialYears)
                {
                    PESA2018Entry pesaEntry = new PESA2018Entry();

                    pesaEntry.FinancialYear = FinancialYear;
                    pesaEntry.YearEnd = int.Parse(FinancialYear.Substring(0, 4)) + 1;
                    pesaEntry.Value = SafeConvertObjectToDouble(CURRENTTABLE.Rows[i].ItemArray[CURRENTTABLE.Rows[0].ItemArray.OfType<string>().ToList().IndexOf(FinancialYear)]);

                    pesaEntry.DepartmentName = (string)CURRENTTABLE.Rows[i].ItemArray[0];
                    pesaEntry.OrganisationName = (string)CURRENTTABLE.Rows[i].ItemArray[1];
                    pesaEntry.CRASegmentCode = (string)CURRENTTABLE.Rows[i].ItemArray[2];
                    pesaEntry.CRASegmentName = (string)CURRENTTABLE.Rows[i].ItemArray[3];
                    pesaEntry.COFOGLevel0 = (string)CURRENTTABLE.Rows[i].ItemArray[4];
                    pesaEntry.HMTFunction = (string)CURRENTTABLE.Rows[i].ItemArray[5];
                    pesaEntry.COFOGLevel1 = (string)CURRENTTABLE.Rows[i].ItemArray[6];
                    pesaEntry.HMTSubfunction = (string)CURRENTTABLE.Rows[i].ItemArray[7];
                    pesaEntry.IDNonID = (string)CURRENTTABLE.Rows[i].ItemArray[8];
                    pesaEntry.CAPorCUR = (string)CURRENTTABLE.Rows[i].ItemArray[9];
                    pesaEntry.CGorLGorPCorBOE = (string)CURRENTTABLE.Rows[i].ItemArray[10];
                    pesaEntry.HMTorDept = (string)CURRENTTABLE.Rows[i].ItemArray[11];
                    pesaEntry.NUTSRegion = (string)CURRENTTABLE.Rows[i].ItemArray[12];
                    pesaEntry.Country = (string)CURRENTTABLE.Rows[i].ItemArray[13];

                    PESA2018Entries.Add(pesaEntry);
                }
            }
            return PESA2018Entries;
        }
            
        static string SafeConvertObjectToString(object input)
        {
            if (input != null && input != System.DBNull.Value && input.GetType() == typeof(string))
            {
                return (string)input;
            }
            else
            {
                return input.ToString();
            }
        }

        static int? SafeConvertObjectToInt(object input)
        {
            if (input != null && input != System.DBNull.Value)
            {
                if (input.GetType() == typeof(string))
                {
                    string inputAsString = (string)input;
                    if (inputAsString.Contains(","))
                    {
                        return int.Parse(inputAsString.Replace(",", ""));
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (input.GetType() == typeof(double))
                {
                    int returnInt = Convert.ToInt32((double)input);
                    return returnInt;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        static double? SafeConvertObjectToDouble(object input)
        {
            if (input != null && input != System.DBNull.Value)
            {
                if (input.GetType() == typeof(string))
                {
                    string inputAsString = (string)input;
                    if (inputAsString.Contains(" - ") && double.TryParse(inputAsString.Substring(0, 1), out double throwaway))
                    {
                        return double.Parse(inputAsString);
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (input.GetType() == typeof(double))
                {
                    double returnDouble = Convert.ToDouble((double)input);
                    return returnDouble;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }



    }

    public class Population
    {
        public int TIME { get; set; }
        public string GEO { get; set; }
        public double Value { get; set; }
    }


    public class Inflation
    {
        public int Year { get; set; }
        public double CPI2015 { get; set; }
    }

    public class StandardSubfunctionName
    {
        public string Subfunction { get; set; }
        public string StandardSubfunction { get; set; }
    }

    public class StandardFunctionName
    {
        public string Function { get; set; }
        public string StandardFunction { get; set; }
    }

    public class StandardGeography
    {
        public string Name { get; set; }
        public string StandardName { get; set; }
        public string GeoCode { get; set;  }
    }

    public class CombinedPESAEntry
    {
        public string FinancialYear { get; set; }
        public int YearEnd { get; set; }
        public string Geography { get; set; }
        public string GeographyCode { get; set; }
        public double? Value { get; set; }
        public string HMTFunction { get; set; }
        public string HMTSubfunction { get; set; }
        public string CAPorCUR { get; set; }
        public double? Value2015PerCapita {get; set;}
        public double CPI2015 { get; set; }
        public double Population { get; set; }
  
    }

    public class PESA2005Entry
    {
        public string TableTitle { get; set; }
        public string FinancialYear { get; set; }
        public string Unit { get; set; }
        public int YearEnd { get; set; }
        public string Geography { get; set; }
        public long? Value { get; set; }
        [Name("HMT Function")]
        public string Function { get; set; }
        [Name("CAP or CUR")]
        public string CAPorCUR { get; set; }
    }

    public class PESA2010Chap9Entry
    {
        [Name("Dept Code")]
        public string DepartmentCode { get; set; }

        [Name("Dept Name")]
        public string DepartmentName { get; set; }
        [Name("COFOG Level 1")]
        public string COFOGLevel1 { get; set; }
        [Name("HMT Functional Classification")]
        public string HMTFunction { get; set; }
        [Name("Programme Object Group")]
        public string ProgrammeObjectGroup { get; set; }

        [Name("Programme Object Group Alias")]
        public string ProgrammeObjectGroupAlias { get; set; }

        [Name("ID or non-ID")]
        public string IDNonID { get; set; }

        [Name("CAP or CUR")]
        public string CAPorCUR { get; set; }
        [Name("CG, LG or PC")]
        public string CGorLGorPCorBOE { get; set; }
        [Name("NUTS 1 region")]
        public string NUTSRegion { get; set; }
        public string FinancialYear { get; set; }
        public int YearEnd { get; set; }
        public double? Value { get; set; }
    }

    public class PESA2010Entry
    {
        [Name("Programme Object Group")]
        public string ProgrammeObjectGroup { get; set; }
        
        [Name("Programme Object Group Alias")]
        public string ProgrammeObjectGroupAlias { get; set; }

        [Name("Dept Name")]
        public string DepartmentName { get; set; }
        [Name("Dept Code")]
        public string DepartmentCode { get; set; }

        [Name("HMT Functional Classification")]
        public string HMTFunction { get; set; }
        [Name("COFOG Level 1")]
        public string COFOGLevel1 { get; set; }
        [Name("COFOG Level 2")]
        public string COFOGLevel2 { get; set; }

        [Name("HMT Sub-function Classification")]
        public string HMTSubfunction { get; set; }

        [Name("ID or non-ID")]
        public string IDNonID { get; set; }

        [Name("CAP or CUR")]
        public string CAPorCUR { get; set; }
        [Name("CG, LG or PC")]
        public string CGorLGorPCorBOE { get; set; }

        [Name("Allocated by HMT or DEPT")]
        public string HMTorDept { get; set; }
        [Name("NUTS 1 region")]
        public string NUTSRegion { get; set; }
        public string FinancialYear { get; set; }
        public int YearEnd { get; set; }
        public double? Value { get; set; }
    }


    public class PESA2014Entry
    {
        [Name("Department")]
        public string DepartmentName { get; set; }

        [Name("CRA Segment")]
        public string CRASegmentCode { get; set; }

        [Name("CRA Segment Description")]
        public string CRASegmentName { get; set; }

        [Name("COFOG Level 0")]
        public string COFOGLevel0 { get; set; }
        [Name("HMT Function")]
        public string HMTFunction { get; set; }
        [Name("COFOG Level 1")]
        public string COFOGLevel1 { get; set; }

        [Name("HMT Subfunction")]
        public string HMTSubfunction { get; set; }

        [Name("ID/non-ID")]
        public string IDNonID { get; set; }

        [Name("CAP or CUR")]
        public string CAPorCUR { get; set; }
        [Name("CG, LG or PC")]
        public string CGorLGorPCorBOE { get; set; }

        [Name("NUTS Region")]
        public string NUTSRegion { get; set; }

        [Name("Country")]
        public string Country { get; set; }
        public string FinancialYear { get; set; }
        public int YearEnd { get; set; }
        public double? Value { get; set; }
    }



    public class PESA2018Entry
    {
        [Name("Department Name")]
        public string DepartmentName { get; set; }
        [Name("Organisation Name")]
        public string OrganisationName { get; set; }
        [Name("CRA Segment Code")]
        public string CRASegmentCode { get; set; }

        [Name("CRA Segment Name")]
        public string CRASegmentName { get; set; }

        [Name("COFOG Level 0")]
        public string COFOGLevel0 { get; set; }
        [Name("HMT Function")]
        public string HMTFunction { get; set; }
        [Name("COFOG Level 1")]
        public string COFOGLevel1 { get; set; }

        [Name("HMT Subfunction")]
        public string HMTSubfunction { get; set; }

        [Name("ID/non-ID")]
        public string IDNonID { get; set; }

        [Name("CAP or CUR")]
        public string CAPorCUR { get; set; }
        [Name("CG, LG, PC, BOE")]
        public string CGorLGorPCorBOE { get; set; }

        [Name("Allocated by HMT or DEPT")]
        public string HMTorDept { get; set; }
        [Name("NUTS Region")]
        public string NUTSRegion { get; set; }

        [Name("Country")]
        public string Country { get; set; }
        public string FinancialYear { get; set; }
        public int YearEnd { get; set; }
        public double? Value { get; set; }
    }
}