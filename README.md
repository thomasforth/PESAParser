# PESAParser
Scripts to parse and combine multiples PESA (public expenditure statistical analyses) and CRA (country and regional accounts) databases and tables into a single file. This gives a full picture of UK public spending using consistent geographies and definitions from 1997/98 (YearEnd 1998) to 2018/19 (YearEnd 2019).

Similar officially reconciled data is available from the ONS as Table S10 in the [Supplementary Tables to the Country and regional public sector finances:](https://www.ons.gov.uk/economy/governmentpublicsectorandtaxes/publicsectorfinance/articles/countryandregionalpublicsectorfinances/financialyearending2018/relateddata) though only higher levels of spending are included here. This work retains data on spending by sub-function.

Code is in C# .NET Core 3. The .sln file will open in Visual Studio 2019 Community Edition and requires a number of NuGet packages for Excel reading and CSV writing.

A [PowerBI explorer of the data is available online](https://app.powerbi.com/view?r=eyJrIjoiZThiNWE2ZDYtZDQ4Ny00YTU4LWExYjItM2JiZDlkNGUwMDBjIiwidCI6IjU3NjE4NTlmLWVlNjMtNDc0ZS04NzQ2LTRkZGNjMGQzZTllNSJ9). The included .pbix file can be opened in PowerBI for further manipulation.

## Why?
This data and analysis powers tools online at https://odileeds.org/projects/jrf/.

## Sources
I download PESA tables and CRA databases from the following URLs,
* Data is available for 1998 to 1999 at
https://webarchive.nationalarchives.gov.uk/20101118111127/http://www.hm-treasury.gov.uk/pespub_pesa03.htm but there is no capital and current split in spending making this data difficult to 
* 2000 to 2005 data is available in Excel table form at 
https://webarchive.nationalarchives.gov.uk/20101128151454/http://www.hm-treasury.gov.uk/pespub_index.htm. We use the 2005 spreadsheet and parse the relevant tables.
* 2005 to 2010 data is available in database form (Excel format) at https://www.gov.uk/government/statistics/country-and-regional-analysis-2010.
* 2008 to 2012 data is available in database form (Excel format) at
https://www.gov.uk/government/statistics/national-statistics-release-october-2012. As far as I know, this is the first release to include HMT Subfunction, capital and current spending separately, and regional breakdowns for England.
* 2010 to 2014 data is available in database form (Excel format) at
https://www.gov.uk/government/statistics/country-and-regional-analysis-2014.
* 2014 to 2018 data is available in database form (Excel format) at 
https://www.gov.uk/government/statistics/country-and-regional-analysis-2018.
* 2015 to 2019 data is available in database form (Excel format) at
https://www.gov.uk/government/statistics/country-and-regional-analysis-2019
Where data is available for the same year more than once it is taken from the most recent source.

Very detailed spending by region (for example, transport broken down into local roads, national roads, local public transport, railways, etc...) is not aggregated into CRA databases before 2014. It can be extracted from table (usually 8.19 or 9.19) of the single year PESA tables linked to from https://webarchive.nationalarchives.gov.uk/20101128151454/http://www.hm-treasury.gov.uk/pespub_index.htm but this does not contain capital and current splits.

Additional data is from the ONS (CPI deflators) and Eurostat (population by year by UK region).

Tables in .csv format are used to standardised names and classifications that change across releases of the PESA tables. Specifically these are `StandardFunctionNames.csv`, `StandardGeographyNames.csv`, and `StandardSubfunctionNames.csv`.

## Output
Output is a single CSV table `ParsedCombinedPESA.csv`. Beyond merging the input tables and standardising names, classifications, and units this includes the population of each region at the given year and the CPI index (2015 = 100) for the given year. The column `Value2015PerCapita` is calculated from this. Cumulative totals for each region, or group of regions as defined in the `GeographyGroupings.xlsx` file.

## License
* I believe that all PESA tables and CRA databases included in this repository are licensed under the [UK Government Open Licence v3](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/).
* Additional data from Eurostat (population by year by UK region) is available under [The Eurostat Data Licence](https://ec.europa.eu/eurostat/about/policies/copyright).
* Additional data from the ONS (CPI deflators) is available under the [UK Government Open Licence v3](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/).
* I license the derived datasets under the UK Government Open License v3 for simplicity, which I believe to be compatible with The Eurostat Data Licence.
* The code is licensed under [The MIT License](https://opensource.org/licenses/MIT).
