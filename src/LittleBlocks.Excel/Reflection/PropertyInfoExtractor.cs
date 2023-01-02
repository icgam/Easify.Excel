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
    public class PropertyInfoExtractor : IPropertyInfoExtractor
    {
        public PropertyInfo GetPropertyInfo<TOwner, TValue>(Expression<Func<TOwner, TValue>> property)
        {
            var expression = property.Body as MemberExpression;
            if (expression == null && property.Body is UnaryExpression)
                expression = ((UnaryExpression) property.Body).Operand as MemberExpression;

            if (expression == null)
                throw new ArgumentException(
                    $"Only Member and Unary expressions supported by the mapper. '{property.Body.GetType().FullName}' " +
                    "is not supported and property metadata can't be extracted from the lambda!");

            var propertyInfo = expression.Member as PropertyInfo;
            if (propertyInfo == null)
                throw new ArgumentException("The lambda expression 'property' should point to a valid Property");
            return propertyInfo;
        }
    }
}