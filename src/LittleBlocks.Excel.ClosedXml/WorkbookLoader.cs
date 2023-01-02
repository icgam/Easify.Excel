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
using System.IO;
using ClosedXML.Excel;
using LittleBlocks.Excel.Exceptions;
using LittleBlocks.Excel.Extensions;
using Microsoft.Extensions.Logging;

namespace LittleBlocks.Excel.ClosedXml
{
    public sealed class WorkbookLoader : IWorkbookLoader
    {
        public WorkbookLoader(ILogger<WorkbookLoader> logger)
        {
            Log = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        private ILogger<WorkbookLoader> Log { get; }

        public IWorkbook Load(string path)
        {
            try
            {
                var fileName = Path.GetFileName(path);
                var stream = FileHelper.ReadFileToMemoryStream(path);
                
                return new Workbook(new XLWorkbook(stream, XLEventTracking.Disabled), fileName);
            }
            catch (FileFormatException ex)
            {
                Log.LogError(ex, "Failed to open '{0}' workbook.", path);
                
                throw new WorkbookIsCorruptedException(path);
            }
        }

        public IWorkbook Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            return new Workbook(new XLWorkbook(stream, XLEventTracking.Disabled));
        }
    }
}