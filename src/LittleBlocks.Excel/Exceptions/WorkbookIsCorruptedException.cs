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

namespace LittleBlocks.Excel.Exceptions
{
    public sealed class WorkbookIsCorruptedException : ExcelException
    {
        public WorkbookIsCorruptedException(string workbook)
            : base(
                GetMessage(workbook))
        {
        }

        private static string GetMessage(string workbook)
        {
            return
                $"Failed to open '{workbook}'. File might be CORRUPT, " +
                "please replace it with valid EXCEL workbook or try open it in EXCEL and save it as 2003 (XLSX) format.";
        }
    }
}