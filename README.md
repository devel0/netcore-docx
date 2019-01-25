# netcore-docx

.NET core docx [**under development**]

[OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK) helper classes

[![devel0 MyGet Build Status](https://www.myget.org/BuildSource/Badge/devel0?identifier=1ae8618a-1f30-46be-a8c5-2cbeb8dc76f1)](https://www.myget.org/)

## install and usage

browse [myget instructions](https://www.myget.org/feed/devel0/package/nuget/netcore-docx)

add `nuget.config` where your solution or csproj that refer this library in order to allow other to restore correcly myget dependencies.

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
