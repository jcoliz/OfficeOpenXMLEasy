# Office Open XML Serializer

This is a .NET Standard 2.1 library designed to make it easy to serialize C# objects from/to Office Open XML spreadsheet documents.

## Background

Sometimes you just need to read and write objects to or from a spreadsheet. Perhaps you're looking for
a simple library to handle all the low-level details required by the [Office Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK).

## Usage

### Namespace

```c#
using jcoliz.OpenOfficeXml.Serializer;
```

### Simple Serialization

```c#
void WriteToSpreadsheet<T>(Stream stream, IEnumerable<T> items) where T: class
{
    using var writer = new SpreadsheetWriter();
    writer.Open(stream);
    writer.Serialize(items);
}
```

### Simple Deserialization

```c#
IEnumerable<T> ReadFromSpreadsheet<T>(Stream stream) where T : class, new()
{
    using var reader = new SpreadsheetReader();
    reader.Open(stream);
    return reader.Deserialize<T>();
}
```

### Sheet Names

Select the sheet name to write into

```c#
writer.Serialize(items, "MySheet");
```

Discover the sheets available in a spreadsheet

```c#
foreach(var sheet in reader.SheetNames)
    Console.WriteLine(sheet);
```

Choose which to deserialize from

```c#
reader.Deserialize<T>("MySheet")
```

### Exclude Properties on Deserialize

You may want to avoid reading in certain properties. For example, I typically don't want Entity Framework IDs
imported from spreadsheets.

```c#
var items = reader.Deserialize<T>(exceptproperties: new string[] { "ID" });
```