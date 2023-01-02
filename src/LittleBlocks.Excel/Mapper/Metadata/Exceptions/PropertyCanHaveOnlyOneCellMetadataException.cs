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

namespace LittleBlocks.Excel.Mapper.Metadata.Exceptions
{
    public sealed class PropertyCanHaveOnlyOneCellMetadataException : ExcelException
    {
        public PropertyCanHaveOnlyOneCellMetadataException(string propertyName, CellMetadata existingMetadata,
            CellMetadata newMetadata)
            : base(GetMesssage(propertyName, existingMetadata, newMetadata))
        {
        }

        private static string GetMesssage(string propertyName, CellMetadata existingMetadata, CellMetadata newMetadata)
        {
            return
                string.Format(
                    "{0} property has its metadata set to '{1}'. Can't apply the new metadata '{2}' to {0} property!",
                    propertyName, existingMetadata, newMetadata);
        }
    }
}