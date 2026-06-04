# Movie Dashboard Project
## Table Content
[Problem Statement](#problem-statement)
[Data Source](#data-source)
[Tools](#tools)
[Data Cleaning](#data-cleaning)
[Dashboard](#dashboard)
[M-Code](#m-code)
[Recommendations](#recommendations)

### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres. 

Your task is to build an Excel dashboard that provides insights into this dataset. The dashboard should help identify:
- The best-performing actors
- The top movies based on box office metrics
- Director performance
- Genre trends
- Seasonal patterns in movie performance
- Any additional insights that can guide future production decisions

The final dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.

### Data Source
Movie Data : The primary dataset used for this analysis is the "Movie_Data_Homework.xlsx" file, containing detailed information about each movie's performance (box office and budget), actors, directors and genres. 
You can download the original datasource here: [Movies Data](https://github.com/user-attachments/files/28590077/Movies_Data_Homework.xlsx)

### Tools
1. Power Query - I used Power Query for Data Cleaning
2. Excel - I used Excel for Data Analysis
3. Pivot Tables - for Creating the dashboard and Visualizations

### Data Cleaning
* Data loading and inspection.
* Handling errors, missing values.
* Data cleaning and formatting. The excel file after the data cleaning & preparation process can be downloaded here - [Movies Data Dashboard](https://github.com/user-attachments/files/28590111/Movies_Data_Dashboard_VS.xlsx)

### Dashboard 
<img width="459" height="353" alt="image" src="https://github.com/user-attachments/assets/fc440e64-16ed-438f-906c-1b44585ec801" />

### M-Code

```
let
    Source = Excel.Workbook(File.Contents("D:\@Desktop\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Promoted headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Merged queries" = Table.NestedJoin(#"Changed Type", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres" = Table.ExpandTableColumn(#"Merged queries", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Removed columns" = Table.RemoveColumns(#"Expanded Genres",{"Column14", "Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21"}),
    #"Reordered columns" = Table.ReorderColumns(#"Removed columns",{"Movie Title", "Release Date", "Wikipedia URL", "Genres.Genre", "Genre_First_ID", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged queries1" = Table.NestedJoin(#"Reordered columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres1" = Table.ExpandTableColumn(#"Merged queries1", "Genres", {"Genre"}, {"Genres.Genre.1"}),
    #"Renamed columns" = Table.RenameColumns(#"Expanded Genres1",{{"Genres.Genre.1", "Genres.Genre.2"}}),
    #"Reordered columns1" = Table.ReorderColumns(#"Renamed columns",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres.Genre.2", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged queries2" = Table.NestedJoin(#"Reordered columns1", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Expanded Directors" = Table.ExpandTableColumn(#"Merged queries2", "Directors", {"Director"}, {"Directors.Director"}),
    #"Reordered columns2" = Table.ReorderColumns(#"Expanded Directors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres.Genre.2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged queries3" = Table.NestedJoin(#"Reordered columns2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors" = Table.ExpandTableColumn(#"Merged queries3", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Reordered columns3" = Table.ReorderColumns(#"Expanded Actors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres.Genre.2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged queries4" = Table.NestedJoin(#"Reordered columns3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors1" = Table.ExpandTableColumn(#"Merged queries4", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered columns4" = Table.ReorderColumns(#"Expanded Actors1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres.Genre.2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors.Actor.1", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged queries5" = Table.NestedJoin(#"Reordered columns4", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors2" = Table.ExpandTableColumn(#"Merged queries5", "Actors", {"Actor"}, {"Actors.Actor.2"}),
    #"Renamed columns1" = Table.RenameColumns(#"Expanded Actors2",{{"Actors.Actor.2", "Actors.Actor.3"}}),
    #"Reordered columns5" = Table.ReorderColumns(#"Renamed columns1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres.Genre.2", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors.Actor.1", "Cast_Third_ID", "Actors.Actor.3", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed columns2" = Table.RenameColumns(#"Reordered columns5",{{"Actors.Actor.1", "Cast_Second_Actor"}, {"Actors.Actor.3", "Cast_Third_Actor"}, {"Actors.Actor", "Cast_First_Actor"}, {"Directors.Director", "Directors Name"}, {"Genres.Genre.2", "Genre_Second_Genre"}, {"Genres.Genre", "Genre_First_Genre"}}),
    #"Merged queries6" = Table.NestedJoin(#"Renamed columns2", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors3" = Table.ExpandTableColumn(#"Merged queries6", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Renamed columns3" = Table.RenameColumns(#"Expanded Actors3",{{"Actors.Actor", "Cast_Fourth_Actor"}}),
    #"Reordered columns6" = Table.ReorderColumns(#"Renamed columns3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First_Genre", "Genre_Second_ID", "Genre_Second_Genre", "Director_First_ID", "Directors Name", "Cast_First_ID", "Cast_First_Actor", "Cast_Second_ID", "Cast_Second_Actor", "Cast_Third_ID", "Cast_Third_Actor", "Cast_Fourth_ID", "Cast_Fourth_Actor", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged queries7" = Table.NestedJoin(#"Reordered columns6", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors4" = Table.ExpandTableColumn(#"Merged queries7", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Renamed columns4" = Table.RenameColumns(#"Expanded Actors4",{{"Actors.Actor", "Cast_Fifth_Actor"}}),
    #"Reordered columns7" = Table.ReorderColumns(#"Renamed columns4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genre_First_Genre", "Genre_Second_ID", "Genre_Second_Genre", "Director_First_ID", "Directors Name", "Cast_First_ID", "Cast_First_Actor", "Cast_Second_ID", "Cast_Second_Actor", "Cast_Third_ID", "Cast_Third_Actor", "Cast_Fourth_ID", "Cast_Fourth_Actor", "Cast_Fifth_ID", "Cast_Fifth_Actor", "Budget ($)", "Box Office Revenue ($)"}),
    #"Added Custom" = Table.AddColumn(#"Reordered columns7", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom",{{"ROI", Percentage.Type}})
in
    #"Changed Type1"
```

### Recommendations
Top 5 genres are Action, Comedy, etc. I would recommend Netflix to produce a movie with one of these genres as they brought in more in box office revenie based on the data from 2012 to 2016

<img width="167" height="81" alt="image" src="https://github.com/user-attachments/assets/0ecb9e8d-03d8-4c0b-b519-33dca7038946" />

Top 5 directors are Chris Renaud, Zack Snyder, etc. These directors have led films with strong financial returns and large audience appeal.

<img width="163" height="77" alt="image" src="https://github.com/user-attachments/assets/7a29c440-1b4d-4e19-ab77-45dbe3b96bca" />


