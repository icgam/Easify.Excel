// This software is part of the LittleBlocks.Excel Library
// Copyright (C) 2018 LittleBlocks
// 
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU Affero General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Affero General Public License for more details.
// 
// You should have received a copy of the GNU Affero General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.
// 

using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using LittleBlocks.Excel.ClosedXml;
using LittleBlocks.Excel.Extensions;
using LittleBlocks.Excel.Mapper;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace LittleBlocks.Excel.SampleApp
{
    internal class Program
    {
        private const string Workbook = "DataSheet_Person_1000_rows.xlsx";

        private static void Main(string[] args)
        {
            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);
            serviceCollection.AddExcel();
            var serviceProvider = serviceCollection.BuildServiceProvider();
            var builder = serviceProvider.GetService<IExcelMapperBuilder>();
            var mapper = builder.Build(LoadWorkbook(Workbook));

            Console.WriteLine("Press any key when ready to load some EXCEL data:)");
            Console.ReadLine();

            var sw = new Stopwatch();

            Console.WriteLine($"Loading '{Workbook}' ...");

            sw.Start();

            var result = mapper.Map<PersonModel>("Sheet1", opt =>
                opt.ForMember(m => m.CustomRowNumber, Resolve.ByValue("Custom Identification No"))).ToList();

            sw.Stop();
            Console.WriteLine($"{result.Count} records loaded from {Workbook} in {sw.ElapsedMilliseconds} milliseconds.");
            Console.WriteLine("Press any key to exit.");
            Console.ReadLine();
        }

        private static void ConfigureServices(IServiceCollection services)
        {
            services.AddLogging(configure => configure.AddConsole());
        }

        private static IWorkbook LoadWorkbook(string dataFile)
        {
            var datafilePath = $"{AppDomain.CurrentDomain.BaseDirectory}\\data\\{dataFile}";
            var workbook = new XLWorkbook(GetMemoryStream(datafilePath), XLEventTracking.Disabled);
            var fileName = Path.GetFileName(datafilePath);
            return new Workbook(workbook, fileName);
        }

        private static Stream GetMemoryStream(string path)
        {
            var bytes = File.ReadAllBytes(path);
            var stream = new MemoryStream();
            stream.Write(bytes, 0, bytes.Length);
            stream.Position = 0;
            return stream;
        }
    }
}
