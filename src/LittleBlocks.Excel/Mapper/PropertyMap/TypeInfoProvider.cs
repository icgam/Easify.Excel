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
using System.Linq;
using System.Reflection;

namespace LittleBlocks.Excel.Mapper.PropertyMap
{
    public class TypeInfoProvider : ITypeInfoProvider
    {
        public IEnumerable<PropertyInfo> GetPublicProperties<TModel>() where TModel : class
        {
            var type = typeof(TModel);
            return
                type.GetProperties()
                    .Where(
                        pi =>
                            pi.CanRead && pi.GetGetMethod() != null && pi.GetGetMethod().IsPublic && pi.CanWrite &&
                            pi.GetSetMethod() != null && pi.GetSetMethod().IsPublic)
                    .ToList();
        }
    }
}