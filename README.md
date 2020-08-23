# A5: Excel integration into dashboard
Consider the data analysis you did in A3. Using either SQL queries, stored procedures, or C# code that stores results to a staging table -

## 1 - build an excel table linked to your database showing the result of your queries:

(name top 5 hot counties, number of cases, case delta, date)

Imported Data

![Image of Tests](/images/window.png)

![Image of Tests](/images/importeddata.png)

Sorted county with delta

![Image of Tests](/images/hot Counties on the last date.png)

``` sql
SELECT TOP (60) MAX(Date) AS LastDay, County, AVG(FIPS) AS Population, MAX(Cases) AS Cases, (MAX(Cases)-Min(Cases)) AS Delta
FROM (
    SELECT *, ROW_NUMBER() OVER(PARTITION BY County ORDER BY Date DESC) AS RowNumber
    FROM dbo.cases5
    WHERE State = 'Washington') AS Temp
WHERE Temp.RowNumber <= 2
GROUP BY County
ORDER BY Delta DESC
```

## 2 - for each hot county, either in a separate sheet or part of sheet, link to the full historical case history, and show a chart of case deltas (new daily cases)

![Image of Tests](/images/eachcounty.png)

``` sql
SELECT *
FROM dbo.cases5
WHERE State = 'Washington' AND County='Yakima'
ORDER BY Date
```
