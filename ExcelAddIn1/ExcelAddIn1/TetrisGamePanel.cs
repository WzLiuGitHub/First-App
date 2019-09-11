using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TetirsGameInstance;
using ExternalCodeNamespace.TetirsGameS;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    using App = Microsoft.Office.Interop.Excel.Application;
    using color = System.Drawing.Color;
    public partial class TetrisGamePanel : Form
    {
        #region field Definition
        public event GameEventHandler RowsClaerd;
        public event GameEventHandler GameFailed;
        public event GameEventHandler GamePaused;
        private Timer GameTimer;
        private Gameblock CurrentBlock;
        private GamePlane Wplane;

        private readonly int RowNum;
        private readonly int ColNum;
        private const int DefStepWise = 10;
        private int Score = 0;
        private Worksheet GameSheet;
        private Random RandomNum;
        private int CurrentType;
        private int NextType;
        private int Gwidth => Wplane.RowWidth;

        private int StepWise;
        #endregion
        #region Static Data Region 
        private static string[] StringShow;
        private static string Heading = "the next block is \r\n \r\n";
        #endregion
        #region Static Data Assignment 
        static TetrisGamePanel()
        {
            StringShow = new string[7];
            StringShow[0] = "       口 口 口\r\n" +
                            "          口   \r\n";

            StringShow[1] = "       口 口\r\n" +
                            "          口 口";

            StringShow[2] = "          口 口\r\n" +
                            "       口 口";

            StringShow[3] = "         口\r\n" +
                            "         口\r\n" +
                            "         口 口";

            StringShow[4] = "           口\r\n" +
                            "           口\r\n" +
                            "        口 口";


            StringShow[5] = "        口\r\n" +
                            "        口\r\n" +
                            "        口\r\n" +
                            "        口";


            StringShow[6] = "      口 口\r\n" +
                            "      口 口";


        }
        #endregion
        
        public TetrisGamePanel(int row , int col , Worksheet ws)
        {
            InitializeComponent();


            if (GameSheet == null)
                GameSheet = ws;
            RowNum = row;
            ColNum = col;
            this.Wplane = new GamePlane(ws,row, col);
            this.GameTimer = new Timer();

            #region Event Define
            GameTimer.Interval = 120;
            {
                Globals.ThisAddIn.Application.WindowResize += Application_WindowResize;
                GameTimer.Tick += GameTimer_Tick;
                Wplane.ClearRows += Wplane_ClearRows;
                Wplane.GameFailed += Wplane_GameFailed;
                
            }
            #endregion

            RandomNum = new Random();

            CurrentType = RandomNum.Next(7);
            NextType = RandomNum.Next(7);
            this.Deactivate += PauseGame;
            StepWise = DefStepWise;
            GameTimer.Enabled = false;
            
            this.CurrentBlock = new Gameblock(ws, col/2*Gwidth, -1*Gwidth, Gwidth,CurrentType, 1);
            textBox1.Text = Heading + StringShow[NextType];
        }
        private int Wplane_ClearRows(int Args)
        {
            Score += Args;
            this.RowsClaerd?.Invoke(Args);
            return Score;
        }
        private int Wplane_GameFailed(int Args)
        {
            MessageBox.Show("You lost and the score is " + Score.ToString());
            foreach (Shape s in this.GameSheet.Shapes)
                s.Delete();
            Wplane = new GamePlane(GameSheet, RowNum, ColNum);
            this.CurrentBlock = new Gameblock(GameSheet, ColNum / 2 * Gwidth, -1 * Gwidth, Gwidth, 1, 1);
            Wplane.Show();
            CurrentBlock.Show();
            return Args;
        }
        private void Application_WindowResize(Workbook Wb, Window Wn)
        {
            if (GameTimer.Enabled)
                GameTimer.Enabled = false;
            Wplane.Show(CurrentBlock.BlockWidth);
            CurrentBlock.SetNewCoordination((CurrentBlock.X/CurrentBlock.BlockWidth)*Gwidth,
                (CurrentBlock.Y/CurrentBlock.BlockWidth)*Gwidth);
            CurrentBlock.Blockwidth = Gwidth;
            CurrentBlock.Show();
        }
        private void GameTimer_Tick(object sender, EventArgs e)
        {
            if (!Wplane.HasPoints(CurrentBlock.DownPoints(StepWise)))
                CurrentBlock.Y += StepWise;
            else
            {
                CurrentBlock.Y = (CurrentBlock.Y / Gwidth + 1)*Gwidth;
                Wplane.AddBlock(CurrentBlock);
                CurrentType = NextType;
                NextType = RandomNum.Next(7);
                CurrentBlock = new Gameblock(GameSheet, ColNum / 2 * Gwidth, -1 * Gwidth, Gwidth, CurrentType, 1);
                textBox1.Text = Heading + StringShow[NextType];
                CurrentBlock.Show();
            }

        }
        private void PauseGame(object sender, EventArgs e)
        {
            GameTimer.Stop();
            GamePaused?.Invoke(0);
        }
        private void TetrisGamePanel_Load(object sender, EventArgs e)
        {
            
        }
        private void ButtonClick(object sender, EventArgs e)
        {
            CurrentBlock.Show();
            Wplane.Show();
            GameTimer.Start();
        }
        private void CloseGame(object sender, FormClosedEventArgs e)
        {
            App app = GameSheet.Application;
            app.DisplayAlerts = false;
            GameSheet?.Delete();
            app.DisplayAlerts = true;
            app.WindowResize -= Application_WindowResize;
        }
        private void ButtonKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.S:
                case Keys.NumPad2:
                    StepWise = DefStepWise * 3;
                    break;
            }
        }
        private void BUttonKeyUp(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.S:
                case Keys.NumPad2:
                    StepWise = DefStepWise;
                    break;
            }
        }
        private void Button1_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            switch (e.KeyChar)
            {
                case 'a':
                case '4':
                    if (!Wplane.HasPoints(CurrentBlock.GetAdjacentBlock(1)))
                        CurrentBlock.X -= Gwidth;
                    break;
                case 'd':
                case '6':
                    if (!Wplane.HasPoints(CurrentBlock.GetAdjacentBlock(0)))
                        CurrentBlock.X += Gwidth;
                    break;
                case 'w':
                case '8':
                    if (!Wplane.HasPoints(CurrentBlock.GetNextDirection()))
                        CurrentBlock.Direction++;
                    break;
            }
        }
    }
    #region ClassDefinition
    class GamePlane : Plane_Base
    {
        private Worksheet Workingsheet;
        private Shape PlaneCanvas;
        private Shape[] VerticalLines;
        private Shape[] HorizonalLines;
        private Shape[][] BlocksShape;
        public event GameEventHandler GameFailed;
        public event GameEventHandler ClearRows;
        public int RowWidth
        {
            get
            {
                double UsableHeight = Workingsheet.Application.UsableHeight;
                double UsableWidth = Workingsheet.Application.UsableWidth;
                double PossibleHeight = UsableHeight < UsableWidth ? UsableHeight : UsableWidth;
                int MaxHeight = (int)PossibleHeight - 50;
                int ColumnWidth = MaxHeight / ColNum;
                int RowWidth = MaxHeight / RowNum;
                return  RowWidth < ColumnWidth ? RowWidth : ColumnWidth;
            }
        }

        public void Show(int OldWidth = 0)
        {
            App PlaneApp = Workingsheet.Application;
            if (RowWidth >= 3)
            {
                PlaneApp.ScreenUpdating = false;
                Worksheet Planesheet = Workingsheet;
                {
                    PlaneCanvas?.Delete();
                    if (VerticalLines != null)
                        foreach (Shape line in VerticalLines)
                            line?.Delete();
                    if (HorizonalLines != null)
                        foreach (Shape line in HorizonalLines)
                            line?.Delete();
                }
                PlaneCanvas = Planesheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle
                    , 0, 0, RowWidth * ColNum, RowWidth * RowNum
                    );
                HorizonalLines = new Shape[RowNum + 1];
                for (int i = 0; i <= RowNum; ++i)
                    HorizonalLines[i] = Planesheet.Shapes.AddLine
                        (0, i * RowWidth, ColNum * RowWidth, i * RowWidth);
                VerticalLines = new Shape[ColNum + 1];
                for (int j = 0; j <= ColNum; ++j)
                    VerticalLines[j] = Planesheet.Shapes.AddLine
                        (j * RowWidth, 0, j * RowWidth, RowNum * RowWidth);
                PlaneCanvas.Fill.ForeColor.RGB = color.White.ToArgb();
                PlaneCanvas.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);


                if (OldWidth != 0)
                {
                    for (int i = 0; i < RowNum; ++i)
                    {
                        if (this.SpaceData[i] != 0)
                        {
                            for (int j = 0; j < ColNum; ++j)
                            {
                                if (BlocksShape[i][j] != null)
                                {
                                    int x = (int)BlocksShape[i][j].Left;
                                    int y = (int)BlocksShape[i][j].Top;
                                    BlocksShape[i][j].Left = x / OldWidth * RowWidth;
                                    BlocksShape[i][j].Top = y / OldWidth * RowWidth;
                                    BlocksShape[i][j].Width = RowWidth;
                                    BlocksShape[i][j].Height = RowWidth;
                                }
                            }
                        }
                    }
                    
                }


                PlaneApp.ScreenUpdating = true;
            }
        }
        public bool HasPoints(SPoint[] points )
        {

            if (points != null)
            {
                foreach (SPoint p in points)
                {
                    if (p.X > ColNum || p.Y > RowNum || p.X <= 0 || p.Y <= 0)
                        return true;
                    if (this[p.Y - 1][p.X - 1] == 1)
                        return true;
                }
                return false;
            }
            throw new Exception();
        }

        public void AddBlock(Gameblock gb)
        {
            if (gb != null)
            {
                App app = this.Workingsheet.Application;
                app.ScreenUpdating = false;
                Shape[] shapes = gb.BlockShape;
                foreach (Shape s in shapes)
                {
                    int x = (int)s.Left;
                    int y = (int)s.Top;
                    this.BlocksShape[y / RowWidth][x / RowWidth] = s;
                    SetValue(y / RowWidth + 1, x / RowWidth + 1, 1);
                }
                int stop = 0;
                int count = 0;
                for (int i = 0; i < RowNum; ++i)
                {
                    
                    if (IsFUllRow(i + 1))
                    {
                        if (stop == 0)
                            stop = i;
                        this.SpaceData[i] = 0;
                        for (int j = 0; j < ColNum; ++j)
                            this.BlocksShape[i][j]?.Delete();
                        count++;
                    }
                    else
                    {
                        if (stop != 0)
                            break;
                    }
                }
                if (count != 0)
                {
                    for (int r = stop + count - 1; r >= 0; --r)
                    {
                        if (r - count < 0)
                        {
                            for (int j = 0; j < ColNum; ++j)
                                this.BlocksShape[r][j] = null;
                        }
                        else
                        {
                            for (int j = 0; j < ColNum; ++j)
                            {
                                this.BlocksShape[r][j] = this.BlocksShape[r - count][j];
                                if (this.BlocksShape[r][j] != null)
                                    BlocksShape[r][j].Top += count * RowWidth;
                            }
                        }
                    }
                    for (int r = stop + count -1; r >= 0; --r)
                    {
                        this.SpaceData[r] = r - count < 0 ? 0 : this.SpaceData[r - count];
                    }

                }


                if (count != 0) ClearRows?.Invoke(count);


                app.ScreenUpdating = true;

                if (this.SpaceData[0] > 0)
                {
                    this.GameFailed?.Invoke(0);
                    return;
                }
            }
        }
        public GamePlane(Worksheet ws , int row, int col) : base(row, col)
        {
            BlocksShape = new Shape[RowNum][];
            for (int i = 0; i < RowNum; ++i)
                BlocksShape[i] = new Shape[ColNum];
            Workingsheet = ws;
        }

    }
    internal class Gameblock : Blocks
    {
        private Worksheet workingsheet;
        public Gameblock(Worksheet ws , int x, int y, int width, int ty, int ori) :
            base(x, y, width, (BlockType)ty, ori)
        {
            BlockShape = new Shape[4];
            workingsheet = ws;
        }
        public Gameblock(Gameblock ts, int x = -1, int y = -1, int width = -1, int ty = -1, int ori = -1)
            : base(x == -1 ? ts.Xcor : x, y == -1 ? ts.Ycor : y, width <= 0 ? ts.Width : width, 
                  (BlockType)(ty < 0 ? ts.Type : ty%7)
                  , ori == -1 ? ts.Oritation : ori) 
        {
            workingsheet = ts.workingsheet;
        }
        public Shape[] BlockShape;

        public void Show()
        {
            
            if (BlockShape != null)
            {
                foreach (Shape shape in BlockShape)
                    shape?.Delete();
            }
            Matrix_Base[] data = GetFullMatrixSet();
            for (int i = 0; i < 4; ++i)
            {
                if (data[i].IsPositive)
                {
                    int top = data[i].GetValue(2, 1);
                    int left = data[i].GetValue(1, 1);
                    BlockShape[i] = workingsheet.Shapes.AddShape
                        (Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                        left, top, Blockwidth, Blockwidth);
                }
                else
                {
                    if (data[i].GetValue(2, 3) > 0)
                        BlockShape[i] = workingsheet.Shapes.AddShape
                        (Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                         data[i].GetValue(1, 1), 0, Blockwidth, data[i].GetValue(2, 4));
                    else
                        BlockShape[i] = null;
                }

            }
            
        }
        public int Blockwidth
        {
            get => this.Width;
            set
            {
                if (value > 0)
                    this.Width = value;
            }
        }
        public int X
        {
            get => Xcor;
            set
            {
                if (value != Xcor)
                {
                    Xcor = value;
                    Show();
                    Changed1 = true;
                }
            }
        }
        public int Y
        {
            get => Ycor;
            set
            {
                if (value != Ycor)
                {
                    Ycor = value;
                    Show();
                    Changed1 = true;
                }
            }
        }
        public int Type
        {
            get => (int)(this.TypeOfBlock);
            set {
                this.TypeOfBlock = (BlockType)(value % 7);
                Show();
            }
        }
        public int Direction
        {
            get => this.Oritation;
            set {
                if (value != Oritation)
                {
                    Changed1 = true;
                    SetNewDirection(value);
                    Show();
                }
            }
        }
        private bool CanBePoints()
        {
            return (Xcor % Blockwidth == 0) && (Ycor % Blockwidth == 0);  
        }
        public SPoint[] DownPoints(int stepwise)
        {
            SPoint[] points;
            SPoint[] Temp = new SPoint[4];
            Matrix_Base[] mx;
            Ycor += stepwise;
            mx = GetFullMatrixSet();
            int count = 0;
            for (int i = 0; i < 4; ++i)
            {
                Temp[i] = new SPoint
                {
                    X = mx[i].GetValue(1, 3) / Blockwidth+1,
                    Y = mx[i].GetValue(2, 3) / Blockwidth+1
                };
                if (Temp[i].Y > 0)
                    ++count;
            }
            points = new SPoint[count];
            int j = 0;
            for (int i = 0; i < 4; ++i)
                if (Temp[i].Y > 0)
                    points[j++] = Temp[i];

            Ycor -= stepwise;
            return points;
        }
        public SPoint[] GetAdjacentBlock(bool IsLeft)
        {

            Xcor += IsLeft ? -Blockwidth : Blockwidth;

            SPoint[] temp;
            Matrix_Base[] data = GetFullMatrixSet();


            if (CanBePoints())
            {
                temp = new SPoint[4];
                List<SPoint> points = new List<SPoint>();
                for (int i = 0; i < 4; ++i)
                {
                    temp[i] = new SPoint
                    {
                        X = data[i].GetValue(1, 1) / Blockwidth + 1,
                        Y = data[i].GetValue(2, 1) / Blockwidth + 1
                    };
                    if (temp[i].Y > 0)
                        points.Add(temp[i]);
                }
                Xcor -= IsLeft ? -Blockwidth : Blockwidth;
                return points.ToArray();
            }



            temp = new SPoint[8];

            for (int i = 0; i < 8; ++i)
                temp[i] = new SPoint
                {
                    X = data[i/2].GetValue(1, i % 2 == 0 ? 1 : 3) / Blockwidth + 1,
                    Y = data[i/2].GetValue(2, i % 2 == 0 ? 1 : 3) / Blockwidth + 1
                };
            List<SPoint> sPoints = new List<SPoint>();
            
            foreach (var p in temp)
                if (!sPoints.Contains(p) && p.Y > 0) 
                    sPoints.Add(p);


            Xcor -= IsLeft ? -Blockwidth : Blockwidth;
            return sPoints.ToArray();
        }
        public SPoint[] GetAdjacentBlock(int Isleft) => GetAdjacentBlock(Isleft != 0);
        public SPoint[] Points
        {
            get
            {
                if (CanBePoints())
                {
                    Matrix_Base[] mx = GetFullMatrixSet();
                    SPoint[] points = new SPoint[4];
                    for (int i = 0; i < 4; ++i)
                        points[i] = new SPoint
                        { X = mx[i].GetValue(1, 1) / Blockwidth+1, Y = mx[i].GetValue(2, 1) / Blockwidth+1 };

                    return points;
                }
                throw new Exception();
            }
        }
        public SPoint[] GetNextDirection()
        {
            SetNewDirection(Oritation + 1);
            Matrix_Base[] mx = GetFullMatrixSet();
            List<SPoint> points = new List<SPoint>();
            if (CanBePoints())
            {
                SPoint[] temp = new SPoint[4];
                for (int i = 0; i < 4; ++i)
                    temp[i] = new SPoint
                    {
                        X = mx[i].GetValue(1, 1) / Width + 1,
                        Y = mx[i].GetValue(2, 1) / Width + 1
                    };
                foreach (SPoint s in temp)
                    if (s.Y > 0)
                        points.Add(s);
            }
            else
            {
                SPoint[] temp = new SPoint[8];
                for (int i = 0; i < 4; ++i)
                {
                    temp[2*i] = new SPoint
                    {
                        X = mx[i].GetValue(1, 1) / Width + 1,
                        Y = mx[i].GetValue(2, 1) / Width + 1
                    };
                    temp[2 * i + 1] = new SPoint
                    {
                        X = mx[i].GetValue(1, 3) / Width + 1,
                        Y = mx[i].GetValue(2, 3) / Width + 1
                    };
                }
                foreach (SPoint s in temp)
                {
                    if (s.Y > 0 && !points.Contains(s))
                        points.Add(s);
                }
            }

            SetNewDirection(Oritation - 1);
            return points.ToArray();
        }
        public bool Changed1 { get; set; } = false;
    }

    #endregion
}
