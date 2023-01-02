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
using LittleBlocks.Excel.Mapper.PropertyMap.Exceptions;
using LittleBlocks.Excel.Mapper.PropertyMap.Selectors;
using LittleBlocks.Excel.Reflection;
using LittleBlocks.Extensions;
using Microsoft.Extensions.Logging;

namespace LittleBlocks.Excel.Mapper.PropertyMap
{
    public class ExcelPropertyMapFactory : IExcelPropertyMapFactory
    {
        private readonly IPropertyNameConvention _conventionsApplier;
        private readonly ITypeInfoProvider _propertiesExtractor;
        private readonly IPropertyInfoExtractor _propertyInfoExtractor;

        public ExcelPropertyMapFactory(ITypeInfoProvider propertiesExtractor,
            IPropertyInfoExtractor propertyInfoExtractor,
            IPropertyNameConvention conventionsApplier,
            ILogger<ExcelPropertyMapFactory> logger)
        {
            _propertiesExtractor = propertiesExtractor ?? throw new ArgumentNullException(nameof(propertiesExtractor));
            _propertyInfoExtractor =
                propertyInfoExtractor ?? throw new ArgumentNullException(nameof(propertyInfoExtractor));
            _conventionsApplier = conventionsApplier ?? throw new ArgumentNullException(nameof(conventionsApplier));
            Log = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        private ILogger<ExcelPropertyMapFactory> Log { get; }

        public IDictionary<string, CellToPropertyMap> CreatePropertyMap<TModel>(IDataSheet worksheet)
            where TModel : class
        {
            return InternalCreatePropertyMap<TModel>(worksheet, null);
        }        
        
        public IDictionary<string, CellToPropertyMap> CreatePropertyMap<TModel>(IDataSheet worksheet, 
            MemberSpec<TModel> specs) where TModel : class
        {
            return InternalCreatePropertyMap(worksheet, specs?.MemberSpecs);
        }

        public IDictionary<string, CellToPropertyMap> CreatePropertyMap<TModel>(IDataSheet worksheet,
            Action<IMemberSpec<TModel>> mappingConfiguration) where TModel : class
        {
            var specs = GetSpecs(mappingConfiguration);
            return InternalCreatePropertyMap(worksheet, specs);
        }        
        
        private IDictionary<string, CellToPropertyMap> InternalCreatePropertyMap<TModel>(IDataSheet worksheet,
            IReadOnlyList<MemberSpec<TModel>> specs) where TModel : class
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            var ensuredSpecs = EnsureSpecs(specs);
            var map = new Dictionary<string, CellToPropertyMap>();
            var properties = _propertiesExtractor.GetPublicProperties<TModel>();
            var headerRow = worksheet.FirstRowUsed();
            var mappingExceptions = new List<MappingException>();

            foreach (var property in properties)
                try
                {
                    var cell = FindMatchingCell(property, headerRow, ensuredSpecs);
                    if (cell.AnyValue())
                    {
                        var cellMap = new CellToPropertyMap(property, cell.ColumnLetter,
                            cell.RowNumber, cell.ToString());
                        map.Add(property.Name, cellMap);
                    }
                }
                catch (MappingException exception)
                {
                    LogFailedToMapProperty(property, exception);
                    mappingExceptions.Add(exception);
                }

            if (mappingExceptions.Any())
                throw new MappingException($"'{mappingExceptions.Count}' mapping issues found.", mappingExceptions);

            return map;
        }

        private IReadOnlyList<MemberSpec<TModel>> EnsureSpecs<TModel>(IReadOnlyList<MemberSpec<TModel>> specs)
        {
            if (specs != null)
                return specs;
            
            var memberSpec = new MemberSpec<TModel>();
            return memberSpec.MemberSpecs;
        }
        
        private IReadOnlyList<MemberSpec<TModel>> GetSpecs<TModel>(Action<IMemberSpec<TModel>> mappingConfiguration)
            where TModel : class
        {
            var memberSpec = new MemberSpec<TModel>();
            if (mappingConfiguration.AnyValue()) mappingConfiguration(memberSpec);
            return memberSpec.MemberSpecs;
        }

        private IDataSheetCell FindMatchingCell<TModel>(PropertyInfo property, IDataSheetRow headerRow,
            IReadOnlyList<MemberSpec<TModel>> specs) where TModel : class
        {
            if (specs.Any())
            {
                var propertySpec =
                    specs.SingleOrDefault(
                        s => _propertyInfoExtractor.GetPropertyInfo(s.DestinationMember).Name == property.Name);
                if (propertySpec.AnyValue())
                {
                    if (propertySpec.Is<IgnoreColumn>())
                    {
                        LogIgnoringProperty(property);
                        return null;
                    }

                    return propertySpec.SelectCellOrThrow(headerRow);
                }
            }

            var selector = Resolve.ByConvention(property.Name, _conventionsApplier,
                StringComparison.InvariantCultureIgnoreCase);
            var cell = selector.SelectCell(headerRow);

            if (cell == null) throw new UnmapppedPropertyException(property.Name, headerRow);

            return cell;
        }

        #region Logs

        private void LogFailedToMapProperty(PropertyInfo property, MappingException exception)
        {
            Log.LogError(exception, "Failed to map {0} property.", property.Name);
        }

        private void LogIgnoringProperty(PropertyInfo property)
        {
            Log.LogWarning("Ignoring {0} property as instructed by mapping configuration.", property.Name);
        }

        #endregion
    }
}