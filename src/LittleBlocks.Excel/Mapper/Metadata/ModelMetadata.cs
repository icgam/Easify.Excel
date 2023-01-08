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
using System.Linq.Expressions;
using System.Reflection;
using LittleBlocks.Excel.Mapper.Metadata.Exceptions;
using LittleBlocks.Excel.Reflection;

namespace LittleBlocks.Excel.Mapper.Metadata
{
    public abstract class ModelMetadata<TModel> where TModel : class
    {
        private readonly Dictionary<string, CellMetadata> _metadataDictionary;
        private readonly IPropertyInfoExtractor _propertyInfoExtractor;

        protected ModelMetadata(IPropertyInfoExtractor propertyInfoExtractor)
        {
            _propertyInfoExtractor = propertyInfoExtractor ?? throw new ArgumentNullException(nameof(propertyInfoExtractor));
            _metadataDictionary = new Dictionary<string, CellMetadata>();
        }

        public void AddCellMetadata(PropertyInfo propertyInfo, string sheetName, string columnLetter,
            int rowNumber)
        {
            if (propertyInfo == null) throw new ArgumentNullException(nameof(propertyInfo));
            AddCellMetadata(propertyInfo.Name, sheetName, columnLetter, rowNumber);
        }

        public void AddCellMetadata(string propertyName, string sheetName, string columnLetter,
            int rowNumber)
        {
            AddCellMetadata(propertyName, string.Empty, sheetName, columnLetter, rowNumber);
        }

        public void AddCellMetadata(PropertyInfo propertyInfo, string fileName, string sheetName, string columnLetter,
            int rowNumber)
        {
            if (propertyInfo == null) throw new ArgumentNullException(nameof(propertyInfo));
            AddCellMetadata(propertyInfo.Name, fileName, sheetName, columnLetter, rowNumber);
        }

        public void AddCellMetadata(string propertyName, string fileName, string sheetName, string columnLetter,
            int rowNumber)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(propertyName));

            var metadataToAdd = new CellMetadata(
                fileName,
                sheetName,
                columnLetter,
                rowNumber
            );

            if (_metadataDictionary.ContainsKey(propertyName))
                throw new PropertyCanHaveOnlyOneCellMetadataException(propertyName, _metadataDictionary[propertyName],
                    metadataToAdd);

            _metadataDictionary.Add(propertyName, metadataToAdd);
        }

        public CellMetadata GetCellMetadata(Expression<Func<TModel, object>> property)
        {
            if (property == null) throw new ArgumentNullException(nameof(property));
            return GetCellMetadata(_propertyInfoExtractor.GetPropertyInfo(property));
        }

        public bool CellHasMetadata(Expression<Func<TModel, object>> property)
        {
            if (property == null) throw new ArgumentNullException(nameof(property));
            return CellHasMetadata(_propertyInfoExtractor.GetPropertyInfo(property));
        }

        public CellMetadata GetCellMetadata(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null) throw new ArgumentNullException(nameof(propertyInfo));
            return GetCellMetadata(propertyInfo.Name);
        }

        public bool CellHasMetadata(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null) throw new ArgumentNullException(nameof(propertyInfo));
            return CellHasMetadata(propertyInfo.Name);
        }

        public CellMetadata GetCellMetadata(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(propertyName));

            if (_metadataDictionary.ContainsKey(propertyName)) return _metadataDictionary[propertyName];

            throw new PropertyMetadataNotFoundException(propertyName);
        }

        public bool CellHasMetadata(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
                throw new ArgumentException("Value cannot be null or whitespace.", nameof(propertyName));
            return _metadataDictionary.ContainsKey(propertyName);
        }
    }
}