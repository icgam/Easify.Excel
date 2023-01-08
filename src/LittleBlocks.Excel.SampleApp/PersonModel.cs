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

namespace LittleBlocks.Excel.SampleApp
{
    internal class PersonModel
    {
        public string Name { get; set; }
        public string PersonName { get; set; }
        public string PersonNameContainsSpecialCharacter { get; set; }
        public decimal? AmountInCurrency { get; set; }
        public string Currency { get; set; }
        public decimal? Salary { get; set; }
        public DateTime? BirthDate { get; set; }
        public string CountryName { get; set; }
        public bool? Certified { get; set; }
        public int? CertificationId { get; set; }
        public int? CustomRowNumber { get; set; }
    }
}