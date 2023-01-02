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
using Microsoft.Extensions.Logging;

namespace LittleBlocks.Excel.Mapper.PropertyMap.Conventions
{
    public class PropertyNamingConventionsBucket : IPropertyNameConvention
    {
        private readonly IEnumerable<IPropertyNameConvention> _conventions;

        public PropertyNamingConventionsBucket(IEnumerable<IPropertyNameConvention> conventions,
            ILogger<PropertyNamingConventionsBucket> logger)
        {
            _conventions = conventions ?? throw new ArgumentNullException(nameof(conventions));
            Log = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        private ILogger<PropertyNamingConventionsBucket> Log { get; }

        public string ApplyConvention(string name)
        {
            LogApplyingConventions(name);

            if (_conventions.Empty())
            {
                LogNoConventionsFounds(name);
                return name;
            }

            LogConventionsFounds();

            foreach (var convention in _conventions)
            {
                LogApplyingConvention(name, convention);
                name = convention.ApplyConvention(name);
                LogConventionApplied(name, convention);
            }

            LogConventionsApplied(name);

            return name;
        }

        #region Logs

        private void LogApplyingConventions(string nameToApplyConvetionTo)
        {
            Log.LogDebug("Applying naming conventions to {0} ...", nameToApplyConvetionTo);
        }

        private void LogNoConventionsFounds(string nameToApplyConvetionTo)
        {
            Log.LogDebug("No naming conventions found. Returned '{0}' as the original value was.",
                nameToApplyConvetionTo);
        }

        private void LogConventionsFounds()
        {
            Log.LogDebug("{0} naming conventions have been found and will be applied ...", _conventions.Count());
        }

        private void LogApplyingConvention(string nameToApplyConvetionTo, IPropertyNameConvention convention)
        {
            Log.LogDebug("Applying {0} naming convention to '{1}' ...", convention.GetType().Name,
                nameToApplyConvetionTo);
        }

        private void LogConventionApplied(string nameToApplyConvetionTo, IPropertyNameConvention convention)
        {
            Log.LogDebug("Naming convention {0} was applied successfully. The resulting name is '{1}'.",
                convention.GetType().Name, nameToApplyConvetionTo);
        }

        private void LogConventionsApplied(string nameToApplyConvetionTo)
        {
            Log.LogDebug("{0} naming conventions have been applied. Resulting property name is '{1}'.",
                _conventions.Count(), nameToApplyConvetionTo);
        }

        #endregion
    }
}