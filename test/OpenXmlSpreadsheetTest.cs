﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace jcoliz.OfficeOpenXml.Serializer.Tests
{
    [TestClass]
    public class SpreadsheetSerializerTests
    {
        public class SimpleItem<T>
        {
            public T Key { get; set; }

            public override bool Equals(object obj)
            {
                return obj is SimpleItem<T> item &&
                    (
                        (Key == null && item.Key == null)
                        ||
                        (Key?.Equals(item.Key) ?? false)
                     );
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(Key);
            }
        }

        public TestContext TestContext { get; set; }

        void WhenWritingToSpreadsheet<T>(Stream stream,IEnumerable<T> items,bool writetodisk = true) where T: class
        {
            {
                using var writer = new SpreadsheetWriter();
                writer.Open(stream);
                writer.Serialize(items, TestContext.TestName);
            }

            stream.Seek(0, SeekOrigin.Begin);

            if (writetodisk)
            {
                var filename = $"Test-{TestContext.TestName}.xlsx";
                File.Delete(filename);
                using var outstream = File.OpenWrite(filename);
                stream.CopyTo(outstream);
                TestContext.AddResultFile(filename);
            }
        }

        private IEnumerable<T> WhenReadAsSpreadsheet<T>(MemoryStream stream, List<string> sheets) where T : class, new()
        {
            stream.Seek(0, SeekOrigin.Begin);
            using var reader = new SpreadsheetReader();
            reader.Open(stream);
            sheets.AddRange(reader.SheetNames);

            return reader.Deserialize<T>(TestContext.TestName);
        }

        public void WriteThenReadBack<T>(IEnumerable<T> items, bool writetodisk = true) where T : class, new()
        {
            // Given: Some items

            // When: Writing it to a spreadsheet using the new methods
            using var stream = new MemoryStream();
            WhenWritingToSpreadsheet(stream, items, writetodisk);

            // And: Reading it back to a spreadsheet
            var sheets = new List<string>();
            var actual = WhenReadAsSpreadsheet<T>(stream, sheets);

            // Then: The spreadsheet is valid, and contains the expected item
            Assert.AreEqual(1, sheets.Count());
            Assert.AreEqual(TestContext.TestName, sheets.Single());
            Assert.IsTrue(actual.SequenceEqual(items));
        }

        [TestMethod]
        public void SimpleWriteString()
        {
            // Given: A very simple string item
            var Items = new List<SimpleItem<string>>() { new SimpleItem<string>() { Key = "Hello, world!" } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteStringNull()
        {
            // Given: A small list of simple string items, one with null key
            var Items = new List<SimpleItem<string>>() { new SimpleItem<string>(), new SimpleItem<string>() { Key = "Hello, world!" } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteDateTime()
        {
            // Given: A very simple item w/ DateTime member
            var Items = new List<SimpleItem<DateTime>>() { new SimpleItem<DateTime>() { Key = new DateTime(2021,06,08) } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteInt32()
        {
            // Given: A very simple item w/ Int32 member
            var Items = new List<SimpleItem<Int32>>() { new SimpleItem<Int32>() { Key = 12345 } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteInt32Nullable()
        {
            // Given: A very simple item w/ nullable int member
            var Items = new List<SimpleItem<int?>>() { new SimpleItem<int?>() { Key = 12345 } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteDecimal()
        {
            // Given: A very simple item w/ decimal member
            var Items = new List<SimpleItem<decimal>>() { new SimpleItem<decimal>() { Key = 123.45m } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteDecimalNullable()
        {
            // Given: A very simple item w/ nullable decimal member
            var Items = new List<SimpleItem<decimal?>>() { new SimpleItem<decimal?>() { Key = 123.45m } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteBoolean()
        {
            // Given: A very simple item w/ boolean member
            var Items = new List<SimpleItem<Boolean>>() { new SimpleItem<Boolean>() { Key = true } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteBooleanNullable()
        {
            // Given: A very simple item w/ boolean member
            var Items = new List<SimpleItem<Boolean?>>() { new SimpleItem<Boolean?>() { Key = true } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        enum TestEnum { Invalid = 0, Good, Bad, Ugly };

        [TestMethod]
        public void SimpleWriteEnum()
        {
            // Given: A very simple item w/ boolean member
            var Items = new List<SimpleItem<TestEnum>>() { new SimpleItem<TestEnum>() { Key = TestEnum.Good } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        public void SimpleWriteEnumNullable()
        {
            // Given: A very simple item w/ boolean member
            var Items = new List<SimpleItem<TestEnum?>>() { new SimpleItem<TestEnum?>() { Key = TestEnum.Good } };

            // When: Writing it to a spreadsheet 
            // And: Reading it back to a spreadsheet
            // Then: The spreadsheet is valid, and contains the expected item
            WriteThenReadBack(Items);
        }

        [TestMethod]
        [ExpectedException(typeof(NotImplementedException))]
        public void CustomColumnNullFails()
        {
            var writer = new SpreadsheetWriter();
            var sheets = writer.SheetNames;
        }

        [TestMethod]
        public void MultipleSheets_Issue1()
        {
            // Given: A spreadsheet with two randomly-named sheets
            var Items1 = new List<SimpleItem<string>>() { new SimpleItem<string>() { Key = "First Sheet1" }, new SimpleItem<string>() { Key = "First Sheet2" } };
            var Items2 = new List<SimpleItem<string>>() { new SimpleItem<string>() { Key = "Second Sheet" } };

            using var stream = new MemoryStream();
            {
                using var writer = new SpreadsheetWriter();
                writer.Open(stream);
                writer.Serialize(Items1, TestContext.TestName + "01");
                writer.Serialize(Items2, TestContext.TestName + "02");
            }

            // (Write it to disk)
            stream.Seek(0, SeekOrigin.Begin);

            var filename = $"Test-{TestContext.TestName}.xlsx";
            File.Delete(filename);
            using var outstream = File.OpenWrite(filename);
            stream.CopyTo(outstream);
            TestContext.AddResultFile(filename);

            // When: Reading it back from a spreadsheet
            var sheets = new List<string>();
            var actual = WhenReadAsSpreadsheet<SimpleItem<string>>(stream, sheets);

            // Then: The items from the first sheet are found
            Assert.AreEqual(2,actual.Count());
            Assert.AreEqual(2,actual.Where(x=>x.Key.StartsWith("First")).Count());
        }

        public class ThirtyMembers
        {
            public int Member_01 { get; set; }
            public int Member_02 { get; set; }
            public int Member_03 { get; set; }
            public int Member_04 { get; set; }
            public int Member_05 { get; set; }
            public int Member_06 { get; set; }
            public int Member_07 { get; set; }
            public int Member_08 { get; set; }
            public int Member_09 { get; set; }
            public int Member_10 { get; set; }
            public int Member_11 { get; set; }
            public int Member_12 { get; set; }
            public int Member_13 { get; set; }
            public int Member_14 { get; set; }
            public int Member_15 { get; set; }
            public int Member_16 { get; set; }
            public int Member_17 { get; set; }
            public int Member_18 { get; set; }
            public int Member_19 { get; set; }
            public int Member_20 { get; set; }
            public int Member_21 { get; set; }
            public int Member_22 { get; set; }
            public int Member_23 { get; set; }
            public int Member_24 { get; set; }
            public int Member_25 { get; set; }
            public int Member_26 { get; set; }
            public int Member_27 { get; set; }
            public int Member_28 { get; set; }
            public int Member_29 { get; set; }
            public int Member_30 { get; set; }

            public override bool Equals(object obj)
            {
                return obj is ThirtyMembers members &&
                       Member_01 == members.Member_01 &&
                       Member_02 == members.Member_02 &&
                       Member_03 == members.Member_03 &&
                       Member_04 == members.Member_04 &&
                       Member_05 == members.Member_05 &&
                       Member_06 == members.Member_06 &&
                       Member_07 == members.Member_07 &&
                       Member_08 == members.Member_08 &&
                       Member_09 == members.Member_09 &&
                       Member_10 == members.Member_10 &&
                       Member_11 == members.Member_11 &&
                       Member_12 == members.Member_12 &&
                       Member_13 == members.Member_13 &&
                       Member_14 == members.Member_14 &&
                       Member_15 == members.Member_15 &&
                       Member_16 == members.Member_16 &&
                       Member_17 == members.Member_17 &&
                       Member_18 == members.Member_18 &&
                       Member_19 == members.Member_19 &&
                       Member_20 == members.Member_20 &&
                       Member_21 == members.Member_21 &&
                       Member_22 == members.Member_22 &&
                       Member_23 == members.Member_23 &&
                       Member_24 == members.Member_24 &&
                       Member_25 == members.Member_25 &&
                       Member_26 == members.Member_26 &&
                       Member_27 == members.Member_27 &&
                       Member_28 == members.Member_28 &&
                       Member_29 == members.Member_29 &&
                       Member_30 == members.Member_30;
            }

            public override int GetHashCode()
            {
                HashCode hash = new HashCode();
                hash.Add(Member_01);
                hash.Add(Member_02);
                hash.Add(Member_03);
                hash.Add(Member_04);
                hash.Add(Member_05);
                hash.Add(Member_06);
                hash.Add(Member_07);
                hash.Add(Member_08);
                hash.Add(Member_09);
                hash.Add(Member_10);
                hash.Add(Member_11);
                hash.Add(Member_12);
                hash.Add(Member_13);
                hash.Add(Member_14);
                hash.Add(Member_15);
                hash.Add(Member_16);
                hash.Add(Member_17);
                hash.Add(Member_18);
                hash.Add(Member_19);
                hash.Add(Member_20);
                hash.Add(Member_21);
                hash.Add(Member_22);
                hash.Add(Member_23);
                hash.Add(Member_24);
                hash.Add(Member_25);
                hash.Add(Member_26);
                hash.Add(Member_27);
                hash.Add(Member_28);
                hash.Add(Member_29);
                hash.Add(Member_30);
                return hash.ToHashCode();
            }
        }

        [TestMethod]
        [ExpectedException(typeof(NotImplementedException))]
        public void SheetNamesFails()
        {
            var writer = new SpreadsheetWriter();
            var sheets = writer.SheetNames;
        }

        [TestMethod]
        public void WriteLongClass()
        {
            var items = new List<ThirtyMembers>()
            {
                new ThirtyMembers() { Member_23 = 23, Member_30 = 30 }
            };

            WriteThenReadBack(items);
        }

        [TestMethod]
        public void ReadWriteEmptyList()
        {
            // Given: A spreadsheet with an empty list
            // When: Reading it in
            WriteThenReadBack(new List<SimpleItem<string>>());

            // Then: Should succeed, because empty list == empty list
        }
    }
}
