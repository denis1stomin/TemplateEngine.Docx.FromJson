# TemplateEngine.Docx.FromJson
`TemplateEngine.Docx.FromJson` is just an experimental app which generates docx documents using data sources in json format.  
`TemplateEngine.Docx.FromJson` uses https://github.com/UNIT6-open/TemplateEngine.Docx under the hood.

## First steps
- `cd <repo root>/src/`
- `dotnet build`
- `cd TemplateEngine.Docx.Runner/`
- `dotnet run -- --template "..\..\templates\Multiple tables.docx" --source "..\..\data-sources\Multiple tables.Data source.Car.json" --output "..\..\_output\doc_car.docx" --finalize --force`
or
- `dotnet run -- --template "..\..\templates\Multiple tables.docx" --source "..\..\data-sources\Multiple tables.Data source.Books.json" --output "..\..\_output\doc_books.docx" --finalize --force`
