
# Movie Dashboard
## Table of Content
[Problem Statement](#problem-statement)
[DataSouce](#data-source)
[Tools](#tools)
[Dashboard](#dashboard)
[M Code](#m-code)
[Recommendations](#recommendations)

### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actor and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres. Your task is to build an Excel dashboard that provides insights into this dataset. The dashboard should help identify: 
- The best-performing actors
- The top movies based on box office metrics
- Directors performance
- Genre trends
- Seasonal patterns in movie performance
- Any additional insights that can guide future production decisions

The dashboard should be clear, interactive, and visionally compelling, enabling Netflix to make data-driven decisions.

### Data Source
Movie Data: The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about each movie's performance (box office and budget), actors, directors, etc

You can download the original data source here:[Movie Dataset Excel file](https://github.com/user-attachments/files/34001159/Movies_Data_Homework.-.original.xlsx)

### Tools
- Power Query - I used Power Query to Clean Data
- Excel - I used Excel for Data Analysis
      - Pivot Tables - for Creating the Dashboards and Visualization

### Data Cleaning
Data loading and inspection.
Handling errors and missing values.
Data cleaning and formatting. The Excel file after the data cleaning and preparation process can be downloaded here:[Movie Dataset HW - Copy.xlsx](https://github.com/user-attachments/files/34001948/Movie.Dataset.HW.-.Copy.xlsx)

### Dashboard
<img width="508" height="353" alt="Screenshot 2026-10-26 233715" src="https://github.com/user-attachments/assets/2aec2a24-2d96-4488-a655-2ddb4365dbe3" />


### M Code
```
let
    Source = Excel.Workbook(File.Contents("C:\Users\jasol\Downloads\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres" = Table.ExpandTableColumn(#"Merged Queries", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Genres",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns",{{"Genres.Genre", "Genre 1"}}),
    #"Merged Queries1" = Table.NestedJoin(#"Renamed Columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres1" = Table.ExpandTableColumn(#"Merged Queries1", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Expanded Genres1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre 1", "Genre_Second_ID", "Genres.Genre", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Reordered Columns1",{{"Genres.Genre", "Genres 2"}}),
    #"Merged Queries2" = Table.NestedJoin(#"Renamed Columns1", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Expanded Directors" = Table.ExpandTableColumn(#"Merged Queries2", "Directors", {"Director"}, {"Directors.Director"}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Expanded Directors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre 1", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns2",{{"Directors.Director", "Directors 1"}}),
    #"Merged Queries3" = Table.NestedJoin(#"Renamed Columns2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors" = Table.ExpandTableColumn(#"Merged Queries3", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Reordered Columns3" = Table.ReorderColumns(#"Expanded Actors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre 1", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors 1", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries4" = Table.NestedJoin(#"Reordered Columns3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors1" = Table.ExpandTableColumn(#"Merged Queries4", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered Columns4" = Table.ReorderColumns(#"Expanded Actors1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre 1", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors 1", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors.Actor.1", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Reordered Columns4",{{"Actors.Actor", "Actors 1"}, {"Actors.Actor.1", "Actors 2"}}),
    #"Merged Queries5" = Table.NestedJoin(#"Renamed Columns3", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors2" = Table.ExpandTableColumn(#"Merged Queries5", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Reordered Columns5" = Table.ReorderColumns(#"Expanded Actors2",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre 1", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors 1", "Cast_First_ID", "Actors 1", "Cast_Second_ID", "Actors 2", "Cast_Third_ID", "Actors.Actor", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Reordered Columns5",{{"Actors.Actor", "Actors 3"}}),
    #"Merged Queries6" = Table.NestedJoin(#"Renamed Columns4", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Renamed Columns5" = Table.RenameColumns(#"Merged Queries6",{{"Actors", "Actors 4"}}),
    #"Expanded Actors 4" = Table.ExpandTableColumn(#"Renamed Columns5", "Actors 4", {"Actor"}, {"Actors 4.Actor"}),
    #"Reordered Columns6" = Table.ReorderColumns(#"Expanded Actors 4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre 1", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors 1", "Cast_First_ID", "Actors 1", "Cast_Second_ID", "Actors 2", "Cast_Third_ID", "Actors 3", "Cast_Fourth_ID", "Actors 4.Actor", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns6" = Table.RenameColumns(#"Reordered Columns6",{{"Actors 4.Actor", "Actors 4"}}),
    #"Merged Queries7" = Table.NestedJoin(#"Renamed Columns6", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors3" = Table.ExpandTableColumn(#"Merged Queries7", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Reordered Columns7" = Table.ReorderColumns(#"Expanded Actors3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre 1", "Genre_Second_ID", "Genres 2", "Director_First_ID", "Directors 1", "Cast_First_ID", "Actors 1", "Cast_Second_ID", "Actors 2", "Cast_Third_ID", "Actors 3", "Cast_Fourth_ID", "Actors 4", "Cast_Fifth_ID", "Actors.Actor", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns7" = Table.RenameColumns(#"Reordered Columns7",{{"Actors.Actor", "Actors 5"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns7", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"ROI", Percentage.Type}})
in
    #"Changed Type1"

```

### Recommendations

Top genres to produce movies that bring the most in box office revenue are Action, Comedy, Drama, Sci-Fi and Adventure

<img width="161" height="59" alt="image" src="https://github.com/user-attachments/assets/743fa521-8cd2-4bca-9f42-5b2713655050" />


Horror Movies produce the best ROI year over year

<img width="245" height="59" alt="image" src="https://github.com/user-attachments/assets/4dd2ea4b-4d43-4826-867a-de8e6e30fcfd" />

Top 5 Performing actors are mentioned below along with box office revenue they produced

<img width="200" height="59" alt="image" src="https://github.com/user-attachments/assets/a8ab30cc-fd2f-4d1b-a188-fc2cf1cd2f57" />

Director Performance below showing directors that produced top 5 box office revenues

<img width="162" height="61" alt="image" src="https://github.com/user-attachments/assets/c4c4de31-6bd8-43ea-8aa0-ba678786e876" />

Seasonal patterns in how a movie performs: the best time to release a movie is May-July and March
Movies that are released in those months bring the most revenue at the box office.

<img width="142" height="94" alt="image" src="https://github.com/user-attachments/assets/17ce321a-ffa7-4cc2-81cf-dd4f959ea97b" />



Based on the information analyzed about movies, actors, directors and genres from 2012 to 2016 the most profitable movie to produce next is a movie in a Horror genre, starring Jennifer Lawrence, Tom Cruise or Hugh Jackman, directed by Chris Renaud, Zack Snyder or Francis Lawrence and released in months of March, May, June or July
