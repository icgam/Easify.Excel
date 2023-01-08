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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace LittleBlocks.Excel.ClosedXml
{
    public class Workbook : IWorkbook
    {
        private readonly XLWorkbook _workbook;

        public Workbook(XLWorkbook workbook, string fileName) : this(workbook)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(fileName));
            FileName = fileName;
        }

        public Workbook(XLWorkbook workbook)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        }

        public string FileName { get; }
        public IWorkbookProperties Properties => new WorkbookProperties(_workbook.Properties);

        public IEnumerable<IDataSheet> Worksheets
        {
            get { return _workbook.Worksheets.Select(w => new DataSheet(w)); }
        }

        public IDataSheet Worksheet(string name)
        {
            if (name == null) throw new ArgumentNullException(nameof(name));

            return !_workbook.TryGetWorksheet(name, out var worksheet) ? null : new DataSheet(worksheet);
        }

        public void Save(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            _workbook.SaveAs(stream);
        }

        public void Save(string filename)
        {
            if (filename == null) throw new ArgumentNullException(nameof(filename));

            _workbook.SaveAs(filename);
        }


        public void Dispose()
        {
            _workbook.Dispose();
        }
    }
}