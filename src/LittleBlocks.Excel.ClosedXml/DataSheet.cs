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
    public class DataSheet : IDataSheet
    {
        private readonly IXLWorksheet _worksheet;

        public DataSheet(IXLWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        public IDataSheetRow FirstRowUsed()
        {
            return new DataSheetRow(_worksheet.FirstRowUsed());
        }        
        
        public IDataSheetRow FirstRow()
        {
            return new DataSheetRow(_worksheet.FirstRow());
        }

        public IDataSheetRow LastRowUsed()
        {
            return new DataSheetRow(_worksheet.LastRowUsed());
        }

        public IEnumerable<IDataSheetRow> Rows(int firstRow, int lastRow)
        {
            if (firstRow <= 0) throw new ArgumentOutOfRangeException(nameof(firstRow));
            if (lastRow <= 0) throw new ArgumentOutOfRangeException(nameof(lastRow));

            return _worksheet.Rows(firstRow, lastRow).Select(r => new DataSheetRow(r));
        }

        public string Name
        {
            get => _worksheet.Name;
            set => _worksheet.Name = value;
        }

        public IDataSheetRange Range(string rangeAddress)
        {
            if (rangeAddress == null) throw new ArgumentNullException(nameof(rangeAddress));

            return new DataSheetRange(_worksheet.Range(rangeAddress));
        }

        public IDataSheetPicture AddPicture(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            
            var image = _worksheet.AddPicture(stream);
            return new DataSheetPicture(image);
        }
        
        public IDataSheetPicture AddPicture(string path)
        {
            if (path == null) throw new ArgumentNullException(nameof(path));

            var image = _worksheet.AddPicture(path);
            return new DataSheetPicture(image);
        }
    }
}