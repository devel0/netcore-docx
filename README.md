# netcore-docx

[![NuGet Badge](https://buildstats.info/nuget/netcore-util)](https://www.nuget.org/packages/netcore-docx/)

.NET core docx

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
