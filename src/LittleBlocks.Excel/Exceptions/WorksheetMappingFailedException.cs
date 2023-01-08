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

using System.Collections.Generic;
using System.Text;

namespace LittleBlocks.Excel.Exceptions
{
    public sealed class WorksheetMappingFailedException : ExcelException
    {
        public WorksheetMappingFailedException(string sheetName, List<ExcelException> exceptions,
            int totalExceptionsCount)
            : base(
                GetMessage(sheetName, exceptions.Count, totalExceptionsCount), ConvertToGeneralExceptions(exceptions))
        {
        }

        private static string GetMessage(string sheetName, int capturedExceptionsCount, int totalExceptionsCount)
        {
            var sb = new StringBuilder();
            sb.AppendLine(
                $"{totalExceptionsCount} errors encountered while mapping '{sheetName}' datasheet. Please investigate the errors for further details:");

            if (totalExceptionsCount > capturedExceptionsCount)
                sb.AppendLine(
                    $"{capturedExceptionsCount} errors limit reached! This usually indicates one or more columns with incorrect mapping.");

            return sb.ToString();
        }
    }
}