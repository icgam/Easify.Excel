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
using ClosedXML.Excel.Drawings;

namespace LittleBlocks.Excel.ClosedXml
{
    public class DataSheetPicture : IDataSheetPicture
    {
        private readonly IXLPicture _picture;

        public DataSheetPicture(IXLPicture picture)
        {
            _picture = picture ?? throw new ArgumentNullException(nameof(picture));
        }
        
        public IDataSheetPicture MoveTo(IDataSheetCell cell)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            var internalCell = ((DataSheetCell) cell).Internal;

            _picture.MoveTo(internalCell);
            return this;
        }        
        
        public IDataSheetPicture MoveTo(int row, int column)
        {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));

            _picture.MoveTo(row, column);
            return this;
        }

        public IDataSheetPicture Scale(double scaleValue)
        {
            if (scaleValue <= 0) throw new ArgumentOutOfRangeException(nameof(scaleValue));
            
            _picture.Scale(scaleValue);
            return this;
        }

        public IDataSheetPicture WithSize(int width, int height)
        {
            _picture.WithSize(width, height);
            return this;
        }
        
        public IDataSheetPicture WithPlacement(PicturePlacement placement)
        {
            _picture.WithPlacement((XLPicturePlacement)placement);
            return this;
        }
    }
}