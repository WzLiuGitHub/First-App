using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using ExternalCodeNamespace.TetirsGameS;
using TetirsGameInstance;
namespace ExcelAddIn1
{
    using GLA = GlobalArg;
    using color = System.Drawing.Color;

    public partial class UserRibbon
    {
        private void BlockTest(object sender, RibbonControlEventArgs e)
        {
            GameBlock gb = new GameBlock(100, 100, 30, Blocks_base.BlockType.Ts, 2);
            CorPoint[] points = gb.GetCorPoints();
            Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            ShowSomething ss = new ShowSomething(ws);
            ss.Show();
        }
        private void DrawPlane(object sender, RibbonControlEventArgs e)
        {
            DrawPlane dp = new DrawPlane(Globals.ThisAddIn.Application.ActiveSheet);
            dp.Show();
        }
        public int RowNum => int.Parse(editBox1.Text);
        public int ColNum => int.Parse(editBox2.Text);

        private void GameBeginButton(object sender, RibbonControlEventArgs e)
        {
            if (RowNum > 6 && ColNum > 4)
            {
                TetrisGamePanel game = new TetrisGamePanel(RowNum, ColNum, Globals.ThisAddIn.Application.Worksheets.Add());

                {
                    game.RowsClaerd += AddCreadit;
                    game.GameFailed += OnFailedGame;
                    game.GamePaused += Game_GamePaused;
                }
                game.Show();
            }
        }

        private int Game_GamePaused(int Args)
        {
            return Args;
        }

        private int OnFailedGame(int Args)
        {
            return Args;
        }

        private int AddCreadit(int args)
        {
            int s = int.Parse(Score.Text) + args;
            this.Score.Text = s.ToString();
            return args;
        }

    }

    public delegate int GameEventHandler(int Args);

    public class GlobalArg
    {
        public static Microsoft.Office.Interop.Excel.Application app;
        public static string[] SplitStringBy(string str, char source)
        {
            string[] result;

            result = str.Split(source);



            int size = 0;
            foreach (string temp in result)
            {
                if (temp.Length != 0
                    &&
                    temp[0] != source)
                    ++size;
            }
            string[] strs;
            if (size != 0)
                strs = new string[size];
            else
            {
                strs = new string[1];
                return strs;
            }
            int strsize = size;
            size = 0;
            for (int i = 0; i < strsize; ++i)
            {
                while (result[size].Length == 0 ||
                    result[size][0] == source)
                    size++;
                strs[i] = result[size];
                size++;
            }



            return strs;
        }
        public static void MakeSquareCells(Range rng, double wid = 0.5)
        {
            Range sin = rng[1][1];
            if (sin == null)
                throw new Exception();
            double WPchar = sin.Width / rng.ColumnWidth;

            rng.ColumnWidth = (wid * 72.0) / WPchar;
            rng.RowHeight = wid * 72.0;
        }
        public static int RGB_generate(int R, int G, int B)
        {

            return 0;
        }
        
    }
}
namespace TetirsGameInstance
{
    struct CorPoint
    {
        public int x;
        public int y;
        public static implicit operator CorPoint(int[] s )
        {
            CorPoint re = new CorPoint
            {
                x = s[0],
                y = s[1]
            };
            return re;
        }
    }
    class GameBlock : Blocks
    {
        public GameBlock(int x, int y, int wid, BlockType ty, int oritation) : base(x, y, wid, ty, oritation) { }
        public CorPoint[] GetCorPoints()
        {
            CorPoint[] data = new CorPoint[4];
            Matrix_Base[] mxs = this.GetFullMatrixSet();
            for (int i = 0; i < 4; ++i)
            {
                int[] Coldata  = mxs[i].GetCol(1);
                data[i] = Coldata;
            }
            return data;
        }
    }

    class GamePlane : Plane_Base
    {
        private Worksheet PlaneSheet;
        private Shape[] VerticalLines;
        private Shape[] HorizonalLines;
        private Shape PlaneCanvas;
        public GamePlane(int row, int col , Worksheet ws) : base(row, col)
        {
            PlaneSheet = ws;
        }
        public bool HaveBlock(int row, int col)
        {
            if (IsLegalPos(row, col))
                return this[row - 1][col - 1] == 1;
            return false;
        }
        public void ClearRow(int row)
        {
            if (IsLegalPos(row, 1))
            {
                this.SpaceData[row - 1] = 0;
            }
        }
        public void RefreshPlane()
        {

        }
    }

}
