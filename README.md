# netcore-docx

[![NuGet Badge](https://buildstats.info/nuget/netcore-docx)](https://www.nuget.org/packages/netcore-docx/)

.NET core docx

- [API Documentation](https://devel0.github.io/netcore-docx/html/annotated.html)
- [Changelog](https://github.com/devel0/netcore-docx/commits/master)

<hr/>

## description

[OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK) helper classes

## install

- [nuget package](https://www.nuget.org/packages/netcore-docx/)

## how this project was built

```sh
mkdir netcore-docx
cd netcore-docx

dotnet new sln
dotnet new classlib -n netcore-docx

cd netcore-docx
dotnet add package DocumentFormat.OpenXml --version 2.8.1
cd ..

dotnet sln netcore-docx.sln add netcore-docx/netcore-docx.csproj
dotnet restore
dotnet build
```

## references

- [open xml sdk doc](https://github.com/OfficeDev/office-content/tree/master/en-us/OpenXMLCon)

