using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace PriceRender
{
    public partial class Form1 : Form
    {
        List<Pr> PrLst = new List<Pr>();
        string _AeProject;
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            NodesGet();
            GenerateScript();
            ApplyScript();
            Render();
        }
        protected bool CheckTime()
        {
            bool RetVal = false;
            richTextBox1.Text = "";

            richTextBox1.Text = "Schedule Times:"+" \n";
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            Application.DoEvents();
            XmlDocument XDoc = new XmlDocument();
            string XmlPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\Nodes.xml";
            if (System.IO.File.Exists(XmlPath))
            {
                XDoc.Load(XmlPath);
                XmlNodeList XmlLst = XDoc.GetElementsByTagName("Schedule");
                for (int i = 0; i < XmlLst.Count; i++)
                {
                    richTextBox1.Text += XmlLst[i].Attributes["Time"].Value.ToString() + " \n";
                    richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    richTextBox1.ScrollToCaret();
                    Application.DoEvents();

                    DateTime CompDate = DateTime.Parse(DateTime.Now.ToShortDateString() + " " + XmlLst[i].Attributes["Time"].Value.ToString());
                    if (DateTime.Now >= CompDate && DateTime.Now <= CompDate.AddMinutes(1))
                    {
                        richTextBox1.Text += XmlLst[i].Attributes["Time"].Value.ToString() + "*** MATCHED *** \n";
                        richTextBox1.SelectionStart = richTextBox1.Text.Length;
                        richTextBox1.ScrollToCaret();
                        Application.DoEvents();
                        RetVal = true;

                    }
                }
            }


            return RetVal;
        }
        protected void GenerateScript()
        {
            try
            {

                _AeProject = ConfigurationSettings.AppSettings["AeProjectPath"].ToString().Trim();

              

                richTextBox1.Text += "Generate Script Started" + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();

                //try
                //{
                //    File.Delete(Path.GetDirectoryName(Application.ExecutablePath) + "//Scr.jsx");
                //}
                //catch
                //{
                //}

                StreamWriter Str = new StreamWriter(Path.GetDirectoryName(Application.ExecutablePath) + "//Scr.jsx");
                Str.WriteLine("app.open(new File(\"" + _AeProject.Replace("\\", "\\\\") + "\"));  ");
                Str.WriteLine("function LayerText(tname,text)  ");
                Str.WriteLine("{  ");
                Str.WriteLine("for(var i = 1; i <= app.project.numItems; i++) {  ");
                Str.WriteLine("var B=app.project.item(i);  ");
                Str.WriteLine("for(var j=1; j <= B.numLayers;j++) {  ");
                Str.WriteLine("	var L=B.layer(j);  ");
                Str.WriteLine("	if(L.name==tname) {  ");
                Str.WriteLine("	L.sourceText.setValue(text);  ");
                Str.WriteLine("	break;  ");
                Str.WriteLine("}  ");
                Str.WriteLine("}  ");
                Str.WriteLine("}  ");
                Str.WriteLine("}  ");


                foreach (Pr item in PrLst)
                {
                    Str.WriteLine(" LayerText (\"" + item.LayerName + "\",\"" + item.Value + "\");  ");
                }


                Str.WriteLine("app.project.save()");
                Str.WriteLine("app.quit();");
                Str.Close();

                richTextBox1.Text += "Generate Script Ended" + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();

            }
            catch (Exception Exp)
            {
                richTextBox1.Text += Exp.Message + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();
            }



        }
        protected void ApplyScript()
        {
            richTextBox1.Text += "Apply Script" + " \n";
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            Application.DoEvents();

            //Process[] localAll = Process.GetProcesses();
            //foreach (Process item in localAll)
            //{
            //    if (item.ProcessName.ToLower() == "afterfx.com")
            //    Thread.Sleep(10000);
            //}

            Process proc = new Process();

            proc.StartInfo.FileName = "\"" + ConfigurationSettings.AppSettings["AeRenderPath"].ToString().Trim() + "afterfx.com" + "\"";

            string ScriptFile = Path.GetDirectoryName(Application.ExecutablePath) + "\\Scr.jsx";
            proc.StartInfo.Arguments = "  -r  " + "\"" + ScriptFile + "\"";
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.EnableRaisingEvents = true;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;
            proc.Start();
            proc.PriorityClass = ProcessPriorityClass.Normal;
            StreamReader reader = proc.StandardOutput;
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                //if (richTextBox1.Lines.Length > 10)
                //{
                //    richTextBox1.Text = "";
                //}
                richTextBox1.Text += (line) + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();

            }
            proc.Close();
        }
        protected void Render()
        {

            label2.Text = DateConversion.GD2JD(DateTime.Now) +" "+ DateTime.Now.ToShortTimeString();

            richTextBox1.Text += "Start Render" + " \n";
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            Application.DoEvents();

            string Date = DateConversion.GD2JD(DateTime.Now);
            Date = Date.Remove(0, 2);
            Process proc = new Process();            

            proc.StartInfo.FileName = "\"" + ConfigurationSettings.AppSettings["AeRenderPath"].ToString().Trim() + "aerender.exe" + "\"";

            string Comp = ConfigurationSettings.AppSettings["AeComposition"].ToString().Trim();

            string DirPathDest = ConfigurationSettings.AppSettings["OutputPath"].ToString().Trim() + "\\" + Date.Replace("\\", "-").Replace("/", "-") + "\\" + ConfigurationSettings.AppSettings["OutputFolderName"].ToString().Trim();
            if (!Directory.Exists(DirPathDest))
                Directory.CreateDirectory(DirPathDest);
            string OutFile = ConfigurationSettings.AppSettings["OutputFilePrefix"].ToString().Trim() + "_" + DateTime.Now.Hour + "_" + DateTime.Now.Minute + "_" + DateTime.Now.Second + ".avi";
            proc.StartInfo.Arguments = " -project " + "\"" + _AeProject + "\"" + "   -comp   \"" + Comp + "\" -output " + "\"" + DirPathDest + "\\" + OutFile + "\"";
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.EnableRaisingEvents = true;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;

            proc.Start();


            proc.PriorityClass = ProcessPriorityClass.Normal;
            StreamReader reader = proc.StandardOutput;
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                if (richTextBox1.Lines.Length > 12)
                {
                    richTextBox1.Text = "";
                }
                richTextBox1.Text += (line) + " \n";
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
                Application.DoEvents();

            }
            proc.Close();

        
            richTextBox1.Text += DateTime.Now.ToString() + " \n";
            richTextBox1.Text += "=======Task Finished======" + " \n";
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            Application.DoEvents();

            timer1.Enabled = true;
        }
        public static Pr GetValue(Pr PrObj)
        {
            SqlCommand sqlCommand = new SqlCommand();
            sqlCommand.CommandText = @"SELECT TOP (1) VAL, DATETIME FROM   STATISTIC_VAL WHERE   (GROUPID = " + PrObj.DbId + ") ORDER BY DATETIME DESC";
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Connection = new SqlConnection(ConfigGetValue("PriceConnectionString", "value"));

            try
            {
                sqlCommand.Connection.Open();
                SqlDataReader Dr = sqlCommand.ExecuteReader();
                while (Dr.Read())
                {
                    PrObj.Value = Dr["val"].ToString();
                    PrObj.DateTime = Dr["datetime"].ToString();
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {

                sqlCommand.Connection.Close();
                sqlCommand.Dispose();
            }
            return PrObj;
        }
        public static string ConfigGetValue(string TagName, string AttrName)
        {
            string RetStr = null;
            XmlDocument XDoc = new XmlDocument();
            string XmlPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\Nodes.xml";
            if (System.IO.File.Exists(XmlPath))
            {
                XDoc.Load(XmlPath);
                XmlNodeList XmlLst = XDoc.GetElementsByTagName(TagName);
                if (XmlLst.Count > 0)
                {
                    RetStr = XmlLst[0].Attributes[AttrName].Value.ToString();
                }
            }
            return RetStr;
        }
        public void NodesGet()
        {
            XmlDocument XDoc = new XmlDocument();
            string XmlPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\Nodes.xml";

            PrLst = new List<Pr>();
            
            if (System.IO.File.Exists(XmlPath))
            {
                XDoc.Load(XmlPath);
                XmlNodeList XmlLst = XDoc.GetElementsByTagName("Layer");
                for (int i = 0; i < XmlLst.Count; i++)
                {
                    Pr PrObj = new Pr();
                    PrObj.DbId = XmlLst[i].Attributes["DbId"].Value.ToString();
                    PrObj.LayerName = XmlLst[i].Attributes["LayerName"].Value.ToString();
                    PrObj.Title = XmlLst[i].Attributes["Title"].Value.ToString();
                    PrObj = GetValue(PrObj);
                    PrLst.Add(PrObj);
                }              
            }
        }
        public class Pr
        {
            public string Title { get; set; }
            public string LayerName { get; set; }
            public string DbId { get; set; }
            public string Value { get; set; }
            public string DateTime { get; set; }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if(CheckTime())
            {
                timer1.Enabled = false;
                button1_Click(null, null);
            }
        }
    }
}
