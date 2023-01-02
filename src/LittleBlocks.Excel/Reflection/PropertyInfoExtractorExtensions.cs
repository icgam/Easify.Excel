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
using System.Linq.Expressions;
using System.Reflection;

namespace LittleBlocks.Excel.Reflection
{
    public static class PropertyInfoExtractorExtensions
    {
        public static PropertyInfo GetPropertyInfo<TOwner, TValue>(this TOwner owner,
            Expression<Func<TOwner, TValue>> property)
        {
            return GetPropertyInfo(property);
        }

        public static PropertyInfo GetPropertyInfo<TOwner, TValue>(Expression<Func<TOwner, TValue>> property)
        {
            var propertyExtractor = new PropertyInfoExtractor();
            return propertyExtractor.GetPropertyInfo(property);
        }

        public static string GetName<TOwner, TValue>(this Expression<Func<TOwner, TValue>> property)
        {
            return GetPropertyInfo(property).Name;
        }

        public static string GetTypeName<TOwner, TValue>(this Expression<Func<TOwner, TValue>> property)
        {
            var type = GetPropertyInfo(property).PropertyType;
            if (type.IsNullable()) type = Nullable.GetUnderlyingType(type);

            return type.Name;
        }

        public static TValue GetValue<TOwner, TValue>(this Expression<Func<TOwner, TValue>> property, TOwner owner)
        {
            return (TValue) GetPropertyInfo(property).GetValue(owner);
        }

        public static bool IsNullable(this Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }
    }
}