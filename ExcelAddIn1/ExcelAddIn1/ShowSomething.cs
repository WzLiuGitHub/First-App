using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExternalCodeNamespace.TetirsGameS;
using TetirsGameInstance;
namespace ExcelAddIn1
{
    public partial class ShowSomething : Form
    {
        private readonly Microsoft.Office.Interop.Excel.Worksheet PlaneSheet;
        private Microsoft.Office.Interop.Excel.Shape[] shapes;
        public ShowSomething(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            PlaneSheet = ws;
            InitializeComponent();
            shapes = null;
            this.FormClosed += ShowSomething_FormClosed;
        }

        private void ShowSomething_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        int X
        {
            get => int.Parse(textBox1.Text);
        }
        int Y => int.Parse(textBox2.Text);
        int Blockwidth => int.Parse(textBox3.Text);
        int Ty => int.Parse(comboBox1.Text)-1;
        int dir => int.Parse(comboBox2.Text);
        private void ShowSomething_Load(object sender, EventArgs e)
        {
            gb = new GameBlock(X, Y, Blockwidth, (Blocks_base.BlockType)Ty, dir);
            this.button2.KeyPress += ShowSomething_KeyPress;
        }

        private void ShowSomething_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (e.KeyChar)
            {
                case 'a':
                case '4':
                    textBox1.Text = (X - 1).ToString();
                    OnFreshClick(null , null);
                    break;
                case 'w':
                case '8':

                    int dirs = dir % gb.MaxDirection;
                    dirs = dirs == (gb.MaxDirection - 1) ? 0 : dirs + 1;
                    comboBox2.Text = dirs.ToString();
                    OnFreshClick(null, null);
                    break;
                case 'd': case '6':
                    textBox1.Text = (X + 1).ToString();
                    OnFreshClick(null, null);
                    break;
                case 's': case '2':
                    textBox2.Text = (Y + 1).ToString();
                    OnFreshClick(null, null);
                    break;
            }
        }
        GameBlock gb;
        private void OnFreshClick(object sender, EventArgs e)
        {
            gb  = new GameBlock(X, Y, Blockwidth,(Blocks_base.BlockType)Ty, dir);
            if (shapes != null)
                foreach (var shap in shapes)
                {
                    shap?.Delete();
                }
            CorPoint[] points = gb.GetCorPoints();
            shapes = new Microsoft.Office.Interop.Excel.Shape[4];
            for (int i = 0; i < 4; ++i)
            {
                shapes[i] = PlaneSheet.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    points[i].x, points[i].y, Blockwidth, Blockwidth
                    );
            }
        }
    }
}
