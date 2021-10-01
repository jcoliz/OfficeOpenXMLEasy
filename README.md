# Office Open XML Easy

This is a .NET Standard 2.1 library designed to make it easy to read and write objects from/to Office Open XML spreadsheet documents

## Background

The [Office Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) is very low-level and picky about getting everything just right. I wanted a simpler inferface for my
applications which only needed to read and write objects from and to a spreadsheet.

## Usage

### Namespace

```c#
using jcoliz.OpenOfficeXml.Easy;
```

### Simple Writing

```c#
void WriteToSpreadsheet<T>(Stream stream, IEnumerable<T> items) where T: class
{
    using var writer = new OpenXmlSpreadsheetWriter();
    writer.Open(stream);
    writer.Write(items);
}
```

### Simple Reading

```c#
IEnumerable<T> ReadFromSpreadsheet<T>(Stream stream) where T : class, new()
{
    using var reader = new OpenXmlSpreadsheetReader();
    reader.Open(stream);
    return reader.Read<T>().ToList();
}
```

### Sheet Names

Select the sheet name to write into

```c#
writer.Write(items, "MySheet");
```

Discover the sheets available in a spreadsheet

```c#
foreach(var sheet in reader.SheetNames)
    Console.WriteLine(sheet);
```

Choose which to read from

```c#
reader.Read<T>("MySheet")
```

### Exclude Properties on Read

You may want to avoid reading in certain properties. For example, I typically don't want my Entity Framework ID's
imported from spreadsheets.

```c#
var items = reader.Read<T>(exceptproperties: new string[] { "ID" });
```