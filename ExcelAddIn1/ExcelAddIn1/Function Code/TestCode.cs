using System;
using System.Collections.Generic;


namespace ExternalCodeNamespace
{
    namespace TetirsGameS
    {
        public class Matrix_Traits<T>
       where T : struct
        {
            public int Rownum { get; }
            public int Colnum { get; }

            private List<T[]> matrix_data;
            public Matrix_Traits(int Rows, int Cols)
            {
                Rows = Rows > 0 ? Rows : -Rows;
                Cols = Cols > 0 ? Cols : -Cols;
                if (Rows * Cols == 0)
                    throw new Exception("Row or Col = 0");
                Rownum = Rows; Colnum = Cols;
                matrix_data = new List<T[]>(Rows);
                T[] temp;
                for (int i = 0; i < Rows; ++i)
                {
                    temp = new T[Cols];
                    matrix_data.Add(temp);
                }
            }
            public Matrix_Traits(T[][] value)
            {
                Rownum = value.Length;
                Colnum = value[0].Length;
                matrix_data = new List<T[]>(Rownum);
                for (int i = 0; i < Rownum; ++i)
                    matrix_data.Add(value[i]);
            }
            public Matrix_Traits(int row, T[] value)
            {
                Rownum = row; Colnum = value.Length / row;
                matrix_data = new List<T[]>(row);
                for (int i = 0; i < row; ++i)
                {
                    T[] ColData = new T[Colnum];
                    for (int j = 0; j < Colnum; ++j)
                        ColData[j] = value[i * Colnum + j];
                    matrix_data.Add(ColData);
                }

            }
            protected T[] this[int i]
            {
                set
                {
                    matrix_data[i] = value;
                }
                get => matrix_data[i];

            }
            public T GetValue(int row, int col)
            {
                return this[Actual_Index(row)][Actual_Index(col)];
            }
            public T[] GetCol(int Col)
            {
                T[] Coldata = new T[Rownum];
                for (int i = 0; i < Rownum; ++i)
                    Coldata[i] = this[i][Actual_Index(Col)];
                return Coldata;
            }
            public T[] GetRow(int Row)
            {
                T[] Rowdata = new T[Colnum];
                for (int i = 0; i < Colnum; ++i)
                    Rowdata[i] = this[Actual_Index(Row)][i];
                return Rowdata;
            }
            public void SetValue(int row, int col, T va)
            {
                this[Actual_Index(row)][Actual_Index(col)] = va;
            }
            protected void SetRowValue(int RowIndex, T[] entireRow)
            {
                for (int i = 0; i < Colnum && i < entireRow.Length; ++i)
                    matrix_data[Actual_Index(RowIndex)][i] = entireRow[i];
            }
            protected void SetColumnValue(int ColumnIndex, T[] entireColumn)
            {
                for (int j = 0; j < Rownum && j < entireColumn.Length; ++j)
                    matrix_data[j][Actual_Index(ColumnIndex)] = entireColumn[j];
            }
            protected bool Addable(Matrix_Traits<T> target)
            {
                return Rownum == target.Rownum && Colnum == target.Colnum;
            }
            protected bool Multipliable(Matrix_Traits<T> traits)
            {
                return Colnum == traits.Rownum;
            }
            private static int Actual_Index(int RowOrCol)
            {
                return RowOrCol - 1;
            }

        }
        public class Matrix_Base : Matrix_Traits<int>
        {
            public Matrix_Base(int row, int col) : base(row, col) { }
            public Matrix_Base(int[][] vs) : base(vs) { }
            public Matrix_Base(int row, int[] v) : base(row, v) { }
            public static Matrix_Base operator +(Matrix_Base x, Matrix_Base y)
            {
                if (!x.Addable(y))
                    return x;
                Matrix_Base result = new Matrix_Base(x.Rownum, x.Colnum);
                for (int i = 0; i < x.Rownum; ++i)
                    for (int j = 0; j < x.Colnum; ++j)
                        result[i][j] = (short)(x[i][j] + y[i][j]);
                return result;
            }
            public static Matrix_Base operator *(Matrix_Base x, int c)
            {
                Matrix_Base result = new Matrix_Base(x.Rownum, x.Colnum);
                for (int i = 0; i < x.Rownum; ++i)
                    for (int j = 0; j < x.Colnum; ++j)
                        result[i][j] = (short)(x[i][j] * c);
                return result;
            }
            public static Matrix_Base operator *(int c, Matrix_Base x) => x * c;
            public static Matrix_Base operator -(Matrix_Base x, Matrix_Base y) => x + (-1 * y);
            public static Matrix_Base operator *(Matrix_Base x, Matrix_Base y)
            {
                if (x.Multipliable(y))
                {
                    Matrix_Base result = new Matrix_Base(x.Rownum, y.Colnum);
                    for (int i = 0; i < x.Rownum; ++i)
                        for (int j = 0; j < y.Colnum; ++j)
                        {
                            for (int k = 0; k < x.Colnum; ++k)
                                result[i][j] += x[i][k] * y[k][j];
                        }
                    return result;
                }
                return x;
            }
            public bool IsPositive
            {
                get
                {
                    for (int i = 0; i < Rownum; ++i)
                        for (int j = 0; j < Colnum; ++j)
                            if (this[i][j] < 0)
                                return false;
                    return true;
                }
            }
        }
        public class Block_Trait
        {
            protected int Xcor;
            protected int Ycor;
            protected int Width;
            protected readonly static Matrix_Base SquareBase;
            public Block_Trait(int X, int Y, int width)
            {
                Xcor = X; Ycor = Y; Width = width;
            }
            protected Matrix_Base BlockMatrix
            {
                get
                {
                    Matrix_Base res = new Matrix_Base(2, 4);
                    int[] Xdata = { 1, 1, 1, 1 };
                    Matrix_Base x = new Matrix_Base(1, Xdata);
                    int[] Ydata = { Xcor, Ycor };
                    Matrix_Base y = new Matrix_Base(2, Ydata);
                    res = (y * x) + SquareBase * Width;
                    return res;
                }
            }
            static Block_Trait()
            {
                int[] Data = { 0, 1, 0, 1, 0, 0, 1, 1 };
                SquareBase = new Matrix_Base(2, Data);
            }
        }
        public class Blocks_base : Block_Trait
        {
            public enum BlockType
            { Ts, Zs, ZsC, Ls, LsC, Is, Os };

            protected int Oritation;
            protected BlockType TypeOfBlock;
            protected static readonly int[][] TypeInformation;
            protected static readonly Matrix_Base[] OritationMatrices;
            public int BlockWidth { get => this.Width; }
            protected void SetWidth(int width)
            {
                Width = width;
            }
            protected void SetBlockType(BlockType type)
            {
                TypeOfBlock = type;
            }
            protected void SetCoordination(int X, int Y)
            {
                this.Xcor = X; this.Ycor = Y;
            }
            protected void SetOritation(int or)
            {
                Oritation = or;
            }
            static Blocks_base()
            {
                #region Data
                int[] data1 = { 1, 123, 032, 103, 120 };
                int[] data2 = { 0, 013, 120 };
                int[] data3 = { 0, 031, 320 };
                int[] data4 = { 0, 032, 103, 210, 321 };
                int[] data5 = { 0, 012, 123, 230, 301 };
                int[] data6 = { 0, 113, 002 };
                int[] data7 = { 2, 123 };
                TypeInformation = new int[7][];
                TypeInformation[0] = data1;
                TypeInformation[1] = data2;
                TypeInformation[2] = data3;
                TypeInformation[3] = data4;
                TypeInformation[4] = data5;
                TypeInformation[5] = data6;
                TypeInformation[6] = data7;

                OritationMatrices = new Matrix_Base[4];
                int[] down = { 0, 0, 0, 0, 1, 1, 1, 1 };
                OritationMatrices[0] = new Matrix_Base(2, down);
                int[] left = { -1, -1, -1, -1, 0, 0, 0, 0 };
                OritationMatrices[1] = new Matrix_Base(2, left);
                int[] up = { 0, 0, 0, 0, -1, -1, -1, -1 };
                OritationMatrices[2] = new Matrix_Base(2, up);
                int[] right = { 1, 1, 1, 1, 0, 0, 0, 0 };
                OritationMatrices[3] = new Matrix_Base(2, right);
                #endregion
            }
            public Blocks_base(int X, int Y, int Width, BlockType type, int Oritation = 0) : base(X, Y, Width)
            {
                TypeOfBlock = type;
                int Typelenth = TypeInformation[(int)type].Length - 1;
                this.Oritation = Oritation % Typelenth;
            }

            private int Ty
            {
                get
                {
                    return (int)TypeOfBlock;
                }
            }
            private int Dir
            { get => Oritation % (TypeInformation[Ty].Length - 1); }
            protected int Max_Dir
            { get => TypeInformation[Ty].Length - 1; }
            protected Matrix_Base[] GetFullMatrixSet()
            {
                Matrix_Base[] result = new Matrix_Base[4];

                result[0] = BlockMatrix;

                int[] BlockData = ExtractTheOritaion();
                switch (TypeOfBlock)
                {
                    case BlockType.Ts:
                        {
                            for (int i = 0; i < 3; ++i)
                                result[i + 1] = result[0] + Width * OritationMatrices[BlockData[i]];
                            break;
                        }
                    case BlockType.Os:
                        {
                            for (int i = 0; i < 3; ++i)
                                result[i + 1] = result[i] + Width * OritationMatrices[BlockData[i]];
                            break;
                        }
                    default:
                        for (int i = 0; i < 2; ++i)
                            result[i + 1] = result[i] + Width * OritationMatrices[BlockData[i]];
                        result[3] = result[0] + Width * OritationMatrices[BlockData[2]];
                        break;
                }

                return result;
            }
            private int[] ExtractTheOritaion()
            {

                int[] data = new int[3];
                int source = TypeInformation[Ty][Dir + 1];
                data[0] = (int)(source / 100);
                source -= data[0] * 100;
                data[1] = source / 10;
                source -= data[1] * 10;
                data[2] = source;
                return data;
            }
        }
        public class Blocks : Blocks_base
        {
            public Blocks(int X, int Y, int width, BlockType ty, int ori = 0) : base(X, Y, width, ty, ori) { }
            public void SetNewDirection(int dir)
            {
                if (dir < 0)
                    dir += MaxDirection;
                this.Oritation = dir % MaxDirection; 
            }
            public void SetNewCoordination(int X, int Y)
            {
                base.SetCoordination(X, Y);
            }
            public int MaxDirection => this.Max_Dir;
        }
        public struct SPoint
        {
            public int X;
            public int Y;
            public static implicit operator SPoint(int[] s)
            {
                SPoint result = new SPoint
                {
                    X = s[0],
                    Y = s.Length > 1 ? s[1] : 0,
                };
                return result;
            }
        }
        public class Plane_Traits
        {
            protected readonly int RowNum;
            protected readonly int ColNum;
            protected readonly int MaxDataInt;

            protected int[] SpaceData;
            protected Plane_Traits(int row, int col)
            {
                RowNum = row; ColNum = col;
                MaxDataInt = (1 << ColNum) - 1;
                if (col > 31)
                    throw new Exception("OverRow");
                SpaceData = new int[row];
            }
            protected int[] this[int row]
            {
                get
                {
                    int[] data = new int[ColNum];
                    for (int i = 0; i < ColNum; ++i)
                        data[i] = (SpaceData[row] >> (ColNum - i - 1)) & 1;
                    return data;
                }
            }
            protected int this[int row, int col]
            {
                get
                {
                    return this[row][col];
                }
            }
            protected void SetValue(int row, int col, int Value)
            {
                int Row = Actual_Row_Col_Num(row);
                int Col = Actual_Row_Col_Num(col);
                Value = Value == 0 ? 0 : 1;
                if (this[Row][Col] != Value)
                {
                    switch (this[Row][Col])
                    {
                        case 0:
                            SpaceData[Row] += 1 << (ColNum - Col-1);
                            break;             
                        case 1:                
                            SpaceData[Row] -= 1 << (ColNum - Col-1);
                            break;
                    }
                }
            }
            private static int Actual_Row_Col_Num(int x)
            {
                x = x > 0 ? x : -x;
                return x - 1;
            }

        }
        public class Plane_Base : Plane_Traits
        {
            //change all public 
            protected Plane_Base(int row, int col) : base(row, col)
            {
            }
            protected bool IsLegalPos(int row, int col)
            {
                return (row >= 1 && col >= 1
                    && row <= this.RowNum && col <= ColNum);
            }
            protected void Setvalue(int row, int col, int S)
            {
                if (IsLegalPos(row, col))
                    SetValue(row, col, S);
            }
            public static Plane_Base operator +(Plane_Base ps, SPoint[] points)
            {
                foreach (var item in points)
                {
                    ps.Setvalue(item.Y, item.X, 1);
                }
                return ps;
            }
            public static Plane_Base operator +(Plane_Base ps, SPoint point)
            {
                ps.Setvalue(point.Y, point.X, 1);
                return ps;
            }

            protected bool IsFUllRow(int row) => SpaceData[row - 1] == MaxDataInt;
            protected bool IsFullCol(int col)
            {
                for (int i = 0; i < RowNum; ++i)
                    if (this[i][col - 1] != 1)
                        return false;
                return true;
            }
            protected bool GetValue(int row, int col)
            {
                if (IsLegalPos(row, col))
                    return this[row - 1][col - 1] == 1;
                return true;
            }
        }

    }

}
