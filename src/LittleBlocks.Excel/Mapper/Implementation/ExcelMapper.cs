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
using System.Reflection;
using LittleBlocks.Excel.Exceptions;
using LittleBlocks.Excel.Mapper.Metadata;
using LittleBlocks.Excel.Mapper.PropertyMap;
using LittleBlocks.Extensions;
using Microsoft.Extensions.Logging;

namespace LittleBlocks.Excel.Mapper.Implementation
{
    public class ExcelMapper : IExcelMapper
    {
        private const int ExceptionsToReportLimit = 100;
        private readonly IModelBuilder _modelBuilder;
        private readonly IExcelPropertyMapFactory _propertyMapFactory;
        private readonly ICellValueConverter _valueConverter;
        private readonly IWorkbook _workbook;

        public ExcelMapper(IWorkbook workbook, IExcelPropertyMapFactory propertyMapFactory,
            ICellValueConverter valueConverter, IModelBuilder modelBuilder,
            ILogger<ExcelMapper> logger)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _propertyMapFactory = propertyMapFactory ?? throw new ArgumentNullException(nameof(propertyMapFactory));
            _valueConverter = valueConverter ?? throw new ArgumentNullException(nameof(valueConverter));
            _modelBuilder = modelBuilder ?? throw new ArgumentNullException(nameof(modelBuilder));
            Log = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        private ILogger<ExcelMapper> Log { get; }

        public IEnumerable<TModel> Map<TModel>(string sheetName) where TModel : class
        {
            return MapInternal<TModel>(sheetName, ds => _propertyMapFactory.CreatePropertyMap(ds, (Action<IMemberSpec<TModel>>)null));
        }

        public IEnumerable<TModel> Map<TModel>(string sheetName, Action<IMemberSpec<TModel>> mappingConfiguration)
            where TModel : class
        {
            if (mappingConfiguration == null) throw new ArgumentNullException(nameof(mappingConfiguration));
            return MapInternal<TModel>(sheetName, ds => _propertyMapFactory.CreatePropertyMap(ds, mappingConfiguration));
        }
        
        public IEnumerable<TModel> Map<TModel>(string sheetName, MemberSpec<TModel> mapperSpec)
            where TModel : class
        {
            if (mapperSpec == null) throw new ArgumentNullException(nameof(mapperSpec));
            return MapInternal<TModel>(sheetName, ds => _propertyMapFactory.CreatePropertyMap(ds, mapperSpec));
        }

        private IEnumerable<TModel> MapInternal<TModel>(string sheetName,
            Func<IDataSheet, IDictionary<string, CellToPropertyMap>> propertyMapper) where TModel : class
        {
            var worksheet = GetWorksheet(sheetName);
            if (worksheet.Empty()) throw new WorksheetNotFoundException(sheetName, _workbook);

            var mappings = propertyMapper(worksheet);
            if (mappings.Empty()) throw new NoMappingsCreatedException<TModel>(sheetName, _workbook.Properties.Title);

            var mappedRows = new List<TModel>();
            var exceptions = new List<ExcelException>();
            var exceptionsCount = 0;
            var rows = GetRowsWithData(worksheet);

            foreach (var rowToMap in rows)
                try
                {
                    var destinationObject = HydrateDestinationObject<TModel>(_workbook.FileName, sheetName, mappings,
                        rowToMap);
                    mappedRows.Add(destinationObject);
                }
                catch (ExcelException exception)
                {
                    LogFailedToMapDatasheet(sheetName, exception);
                    exceptionsCount++;
                    if (exceptionsCount <= ExceptionsToReportLimit) exceptions.Add(exception);
                }

            if (exceptions.Any()) throw new WorksheetMappingFailedException(sheetName, exceptions, exceptionsCount);

            return mappedRows;
        }

        private IDataSheet GetWorksheet(string sheetName)
        {
            var worksheet =
                _workbook.Worksheets.SingleOrDefault(
                    w => string.Compare(w.Name, sheetName, StringComparison.InvariantCultureIgnoreCase) == 0);
            return worksheet;
        }

        private IEnumerable<IDataSheetRow> GetRowsWithData(IDataSheet worksheet)
        {
            var firstRow = worksheet.FirstRowUsed().RowNumber + 1;
            var lastRow = worksheet.LastRowUsed().RowNumber;
            var rows = worksheet.Rows(firstRow, lastRow);
            return rows;
        }

        private TModel HydrateDestinationObject<TModel>(string fileName, string sheetName,
            IDictionary<string, CellToPropertyMap> mappings, IDataSheetRow rowToMap)
            where TModel : class
        {
            var destinationObject = _modelBuilder.Build<TModel>();

            foreach (var cellToPropertyMapping in mappings)
            {
                var mapping = cellToPropertyMapping.Value;
                var cell = rowToMap.Cell(mapping.CellColumn);
                var convertedValue = _valueConverter.ConvertTo(cell, mapping.Property.PropertyType);
                mapping.Property.SetValue(destinationObject, convertedValue);
                destinationObject = AddPropertyMetadataIfPossible(mapping.Property, fileName, sheetName, cell,
                    destinationObject);
            }

            return destinationObject;
        }

        private static TModel AddPropertyMetadataIfPossible<TModel>(PropertyInfo property, string fileName, string sheetName,
            IDataSheetCellCoordinates cell,
            TModel model) where TModel : class
        {
            var metadata = model as ModelMetadata<TModel>;
            if (!metadata.AnyValue()) 
                return model;
            
            var columnLetter = cell.ColumnLetter;
            var rowNumber = cell.RowNumber;
            metadata.AddCellMetadata(property, fileName, sheetName, columnLetter, rowNumber);

            return model;
        }

        #region Logs

        private void LogFailedToMapDatasheet(string sheetName, ExcelException exception)
        {
            Log.LogError(exception, "Failed to map {0} datasheet. Please consult returned error descriptions.",
                sheetName);
        }

        #endregion
    }
}