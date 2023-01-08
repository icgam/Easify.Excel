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
using LittleBlocks.Excel.Exceptions;

namespace LittleBlocks.Excel.Extensions
{
    public static class WorkbookExtensions
    {
        public static MemoryStream ToMemoryStream(this IWorkbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            
            try
            {
                var stream = new MemoryStream();
                workbook.Save(stream);
                stream.Position = 0;

                return stream;
            }
            catch (Exception e)
            {
                throw new WorkbookToStreamException(e);
            }
        }

        public static IDataSheetRange GetDataSheetRange(this IWorkbook workbook, string sheetName, string range)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
            if (range == null) throw new ArgumentNullException(nameof(range));
            
            var worksheet = workbook.Worksheet(sheetName);
            return worksheet?.Range(range);
        }

        public static void SaveInDataSheetRange(this IWorkbook workbook, string sheetName, string rangeName, List<object[]> data)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
            if (rangeName == null) throw new ArgumentNullException(nameof(rangeName));
            if (data == null) throw new ArgumentNullException(nameof(data));
            
            var worksheet = workbook.Worksheet(sheetName);
            var range = worksheet?.Range(rangeName);
            if (range == null)
                return;

            range.Cell(1, 1).Value = data;
        }        

        public static void SaveInDataSheetRange(this IWorkbook workbook, string sheetName, string rangeName, object data)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
            if (rangeName == null) throw new ArgumentNullException(nameof(rangeName));
            if (data == null) throw new ArgumentNullException(nameof(data));
            
            var worksheet = workbook.Worksheet(sheetName);
            var range = worksheet?.Range(rangeName);
            if (range == null)
                return;

            range.Cell(1, 1).Value = data;
        }
        
        public static void SaveInDataSheet(this IWorkbook workbook, string sheetName, List<object[]> data)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
            if (data == null) throw new ArgumentNullException(nameof(data));
            
            var worksheet = workbook.Worksheet(sheetName);
            if (worksheet == null)
                return;
            
            worksheet.FirstRow().Cell(1).Value = data;
        }


        public static void HideBlankRowsInDataSheet(this IWorkbook workbook, string sheetName, string rangeName)
        {
            const string cellColumn = "1";

            var worksheet = workbook.Worksheet(sheetName);

            var range = worksheet?.Range(rangeName);
            if (range == null)
                return;

            for (var i = 1; i <= range.RowCount; i++)
            {
                var row = range.Row(i);
                var cell = row.Cell(cellColumn);

                if (cell.IsEmpty() || string.IsNullOrWhiteSpace(cell.GetString())) cell.Row().Hide();
            }
        }
    }
}