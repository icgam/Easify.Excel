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
using LittleBlocks.Excel.Mapper.PropertyMap;
using LittleBlocks.Excel.Mapper.PropertyMap.Conventions;
using LittleBlocks.Excel.Reflection;
using LittleBlocks.Extensions;
using Microsoft.Extensions.Logging;

namespace LittleBlocks.Excel.Mapper.Implementation
{
    public class ExcelMapperBuilder : IExcelMapperBuilder
    {
        private readonly ILoggerFactory _loggerFactory;
        private readonly IWorkbookLoader _workbookLoader;

        public ExcelMapperBuilder(IWorkbookLoader workbookLoader, ILoggerFactory loggerFactory)
        {
            _workbookLoader = workbookLoader ?? throw new ArgumentNullException(nameof(workbookLoader));
            _loggerFactory = loggerFactory ?? throw new ArgumentNullException(nameof(loggerFactory));
        }

        public IExcelMapper Build(IWorkbook workbook)
        {
            return Build(workbook, new MapperSettings());
        }

        public IExcelMapper Build(IWorkbook workbook, MapperSettings settings)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            IConverterProvider converterProvider = GetDefaultConvertersProvider();
            IPropertyNameConvention propertyNameConvention = GetDefaultPropertyNamingConventions();

            if (settings.Convention.AnyValue()) propertyNameConvention = settings.Convention;

            if (settings.ConverterProvider.AnyValue()) converterProvider = settings.ConverterProvider;

            var propertiesExtractor = new TypeInfoProvider();
            var propertyInfoExtractor = new PropertyInfoExtractor();
            var propertiesMapFactory = new ExcelPropertyMapFactory(propertiesExtractor, propertyInfoExtractor,
                propertyNameConvention, _loggerFactory.CreateLogger<ExcelPropertyMapFactory>());

            var typeConverter = new CellValueConverter(converterProvider);
            var modelFactory = new ModelBuilder(propertyInfoExtractor);
            return new ExcelMapper(workbook, propertiesMapFactory, typeConverter, modelFactory,
                _loggerFactory.CreateLogger<ExcelMapper>());
        }

        public IExcelMapper Build(string path)
        {
            return Build(path, new MapperSettings());
        }

        public IExcelMapper Build(string path, MapperSettings settings)
        {
            if (path == null) throw new ArgumentNullException(nameof(path));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            var workbook = _workbookLoader.Load(path);
            return Build(workbook, settings);
        }

        private HardcodedConverterProvider GetDefaultConvertersProvider()
        {
            return new HardcodedConverterProvider();
        }

        private PropertyNamingConventionsBucket GetDefaultPropertyNamingConventions()
        {
            var conventions = new PropertyNamingConventionsBucket(new List<IPropertyNameConvention>
            {
                new NameWithNoNumberAtStartConvention(),
                new ReplaceAmperstandWithNConvention(),
                new RemoveNonLetterAndDigitCharactersConvention()
            }, _loggerFactory.CreateLogger<PropertyNamingConventionsBucket>());

            return conventions;
        }
    }
}