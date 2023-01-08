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

namespace LittleBlocks.Excel.Mapper.PropertyMap
{
    public class MemberSpec<TDestination> : IMemberSpec<TDestination>
    {
        private readonly List<MemberSpec<TDestination>> _memberSpecs;
        private IColumnSelector _columnSelector;

        public MemberSpec()
        {
            _memberSpecs = new List<MemberSpec<TDestination>>();
        }

        public MemberSpec(List<MemberSpec<TDestination>> memberSpecs)
        {
            if (memberSpecs == null) throw new ArgumentNullException(nameof(memberSpecs));
            _memberSpecs = memberSpecs;
        }

        public Expression<Func<TDestination, object>> DestinationMember { get; private set; }

        public IReadOnlyList<MemberSpec<TDestination>> MemberSpecs => _memberSpecs;

        public IDataSheetCell SelectCellOrThrow(IDataSheetRow headerRow)
        {
            return _columnSelector.SelectCellOrThrow(headerRow);
        }

        public IDataSheetCell SelectCell(IDataSheetRow headerRow)
        {
            return _columnSelector.SelectCell(headerRow);
        }

        public bool Is<TSelector>() where TSelector : class, IColumnSelector
        {
            return _columnSelector.Is<TSelector>();
        }

        public IMemberSpec<TDestination> ForMember(Expression<Func<TDestination, object>> destinationMember,
            IColumnSelector selector)
        {
            if (destinationMember == null) throw new ArgumentNullException(nameof(destinationMember));
            if (selector == null) throw new ArgumentNullException(nameof(selector));

            DestinationMember = destinationMember;
            _columnSelector = selector;

            _memberSpecs.Add(this);

            return new MemberSpec<TDestination>(_memberSpecs);
        }
    }
}