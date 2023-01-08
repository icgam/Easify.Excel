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
using System.Linq;
using LittleBlocks.Extensions;

namespace LittleBlocks.Excel
{
    public abstract class ExcelException : Exception
    {
        protected ExcelException(string message) : base(message)
        {
        }

        protected ExcelException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        protected ExcelException(string message, IEnumerable<Exception> innerExceptions)
            : base(message, CreateAggregate(message, innerExceptions))
        {
        }

        public IReadOnlyCollection<Exception> Exceptions
        {
            get
            {
                if (InnerException is AggregateException exception)
                    return exception.InnerExceptions.ToList();

                if (InnerException.AnyValue())
                    return new List<Exception>
                    {
                        InnerException
                    };

                return new List<Exception>
                {
                    this
                };
            }
        }

        private static Exception CreateAggregate(string message, IEnumerable<Exception> innerExceptions)
        {
            return new AggregateException(message, innerExceptions);
        }

        protected static IEnumerable<Exception> ConvertToGeneralExceptions<TException>(
            IEnumerable<TException> innerExceptions)
            where TException : ExcelException
        {
            return innerExceptions.Select(e => (Exception) e).ToList();
        }
    }
}