using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using TetirsGameInstance;
using ExternalCodeNamespace.TetirsGameS;
namespace ExcelAddIn1
{
    using color = System.Drawing.Color;
    public partial class DrawPlane : Form
    {


        private int RowNum => int.Parse(textBox1.Text);
        private int ColNum => int.Parse(textBox2.Text);
        private Microsoft.Office.Interop.Excel.Application PlaneApp;
        private int MaxHeight
        {
            get
            {
                double PossibleHeight = PlaneApp.UsableHeight < PlaneApp.UsableWidth ? PlaneApp.UsableHeight : PlaneApp.UsableWidth;
                return   (int)PossibleHeight - 50;
            }
        }
        private Worksheet Planesheet;
        private Plane_Base PlaneData;
        private Shape PlaneCanvas;
        private Shape[] VerticalLines;
        private Shape[] HorizonalLines;
        public DrawPlane(Worksheet planesheet )
        {
            this.Planesheet = planesheet;
            PlaneApp = Planesheet.Application;
            
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void RefreshPlane(object sender, EventArgs e)
        {

            int ColumnWidth = MaxHeight / ColNum;
            int RowWidth = MaxHeight / RowNum;
            RowWidth = RowWidth < ColumnWidth ? RowWidth : ColumnWidth;

            if (RowWidth >= 1 && MaxHeight >= 10)
            {
                PlaneApp.ScreenUpdating = false;
               
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
                HorizonalLines = new Shape[RowNum +1];
                for (int i = 0; i <= RowNum; ++i)
                    HorizonalLines[i] = Planesheet.Shapes.AddLine
                        (0, i * RowWidth, ColNum * RowWidth, i * RowWidth);
                VerticalLines = new Shape[ColNum + 1];
                for (int j = 0; j <= ColNum; ++j)
                    VerticalLines[j] = Planesheet.Shapes.AddLine
                        (j*RowWidth , 0 , j*RowWidth , RowNum*RowWidth);
                PlaneCanvas.Fill.ForeColor.RGB = color.White.ToArgb();
                PlaneCanvas.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);

                PlaneApp.ScreenUpdating = true;
            }

        }

        private void DrawPlane_Load(object sender, EventArgs e)
        {
            PlaneApp.WindowResize += RefreshEvent;
        }

        private void RefreshEvent(Workbook Wb, Window Wn)
        {
            RefreshPlane(null, null);
        }

        private void DrawPlane_FormClosed(object sender, FormClosedEventArgs e)
        {
            
            this.PlaneApp.WindowResize -= RefreshEvent;
        }

        private void Button2_KeyDown(object sender, KeyEventArgs e)
        {
            this.textBox1.Text = "Key down " + e.KeyValue;

            this.textBox2.Text = "10";
        }

        private void Button2_KeyUp(object sender, KeyEventArgs e)
        {
            this.textBox1.Text = "Key up " + e.KeyValue;
            this.textBox2.Text = "30";
        }

        private void Button2_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
    }
}
