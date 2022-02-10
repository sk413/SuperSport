using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;
using Part = Tekla.Structures.Model.Part;

namespace NPTFCheck
{
    public partial class Form1 : Form
    {
        public static string skApplicationName = "NPTF Text finder in Drawing ";
        public static string skApplicationVersion = "2201.02";  

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            skWinLib.accesslog(skApplicationName, skApplicationVersion, "Create Report", "");
            int ct = 0;

            DrawingHandler drawingHandler = new DrawingHandler();
            List<string> missingNPTF = new List<string>();
            List<string> UnwantedNPTF = new List<string>();
            List<string> catchFile = new List<string>();
            DrawingEnumerator drawingEnumerator = drawingHandler.GetDrawingSelector().GetSelected();
            label1.Text = "Processing.........";
            label1.Visible = true;
            int scount = drawingEnumerator.GetSize();
            int a = 0;
            label4.Text = scount.ToString();
            label4.Visible = true;
            label2.Visible = true;
            while(drawingEnumerator.MoveNext())
            {
                ct++;
                AssemblyDrawing assemblyDrawing = drawingEnumerator.Current as AssemblyDrawing;
                try
                {
                    if (assemblyDrawing != null)
                    {
                        Tekla.Structures.Model.ModelObject modelObject = new Model().SelectModelObject(assemblyDrawing.AssemblyIdentifier);
                        Assembly assembly = modelObject as Assembly;
                        Tekla.Structures.Model.Part part = assembly.GetMainPart() as Part;
                        string studs = "";
                        part.GetUserProperty("USER_FIELD_1", ref studs);
                        if(studs.Trim(' ') =="")
                        {
                            int stud = 0;
                            part.GetUserProperty("SHEARSTUDS", ref stud);
                            studs = stud.ToString();

                        }
                        DrawingObjectEnumerator drawingObjectEnumerator = assemblyDrawing.GetSheet().GetAllObjects(typeof(Text));
                        bool containsPaint = false;
                        while (drawingObjectEnumerator.MoveNext())
                        {
                            ct++;
                            Text text = drawingObjectEnumerator.Current as Text;
                            string Mytext = text.TextString.ToUpper().Trim().Replace(" ","");
                                if (Mytext.Equals("NOPAINT\n@TOPFLG."))
                                {
                                    containsPaint = true;
                                    break;
                                }
                                else if (Mytext.Equals("NPTF"))
                                {
                                    containsPaint = true;
                                    break;
                                }

                        }
                        if ((studs !="0") && (containsPaint == false))
                        {
                            missingNPTF.Add(assemblyDrawing.Mark);
                        }
                        else if ((studs =="0") && (containsPaint == true))
                        {
                            UnwantedNPTF.Add(assemblyDrawing.Mark);
                        }
                    }
                }
                catch
                {
                    catchFile.Add(assemblyDrawing.Mark);
                }
                a++;


                label2.Text = a.ToString();
            }

            createReport(missingNPTF,"MissingNPTF");
            createReport(catchFile, "CatchFile");
            createReport(UnwantedNPTF, "UnwantedNPTF");
            label1.Visible = false;
            MessageBox.Show("Completed");

            skWinLib.worklog(skApplicationName, skApplicationVersion, "Create Report", ";Total:" + ct.ToString());
        }
        private void createReport(List<string> drawingNumber,string fileName)
        {
            Model model = new Model();
            ModelInfo modelinfo = model.GetInfo();
            System.IO.Directory.CreateDirectory(modelinfo.ModelPath + "\\Automation");
            String reportPath = modelinfo.ModelPath + "\\Automation\\"+fileName+".csv";
            StreamWriter sw = new StreamWriter(reportPath, false);
            foreach (string reqInfo in drawingNumber)
            {
                sw.Write(reqInfo);
                sw.WriteLine();
            }
            sw.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = chkontop.Checked;
        }
    }
}
