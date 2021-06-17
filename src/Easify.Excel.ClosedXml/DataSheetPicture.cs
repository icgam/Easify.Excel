// This software is part of the Easify.Excel Library
// Copyright (C) 2018 Intermediate Capital Group
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

namespace Easify.Excel.ClosedXml
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
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
            
            _picture.WithSize(width, height);
            return this;
        }
        
        public IDataSheetPicture WithWidth(int width)
        {
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            
            var height = width/(decimal)_picture.Width * _picture.Height;
            _picture.WithSize(width, (int)height);
            return this;
        }        
        
        public IDataSheetPicture WithHeight(int height)
        {
            if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
            
            var width = (height/(decimal)_picture.Height) * _picture.Width;
            _picture.WithSize((int)width, height);
            return this;
        }

        public int Width => _picture.Width;
        public int Height => _picture.Height;
    }
}