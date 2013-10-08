namespace ActiveScriptTest
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;
    using ActiveXScriptLib;
    using System.Runtime.InteropServices;

    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            string addFunc =
                "Public Function Add(a, b) " + Environment.NewLine +
                    "   Add = a + b " + Environment.NewLine +
                "End Function";

            string subtract =
                "Public Function Subtract(a, b) " + Environment.NewLine +
                    "   Subtract = a - b " + Environment.NewLine +
                "End Function";

            string echo =
                "Public Sub SayHello() " + Environment.NewLine + 
                "   WScript.Echo \"Hello World\" " + Environment.NewLine + 
                "End Sub";

            string addCode =
                "Public Sub AddCode() " + Environment.NewLine +
                "   Import \"MyFile.vbs\" " + Environment.NewLine +
                "End Sub";

            try
            {
                ActiveScriptEngine scriptEngine = new ActiveScriptEngine("VBScript");

                scriptEngine.AddCode(addFunc, "Math");
                scriptEngine.AddCode(subtract, "Math");

                scriptEngine.AddCode(
                    "Public Function Add(a, b) " + Environment.NewLine +
                    "   Add = Math.Add(a, b) " + Environment.NewLine +
                    "End Function");

                scriptEngine.AddCode(echo);

                scriptEngine.Start();

                dynamic script = scriptEngine.GetScriptHandle();

                MessageBox.Show(script.Math.Subtract(script.Add(1, 3), 2).ToString());

                scriptEngine.AddObject("WScript", new HostObject());

                script.SayHello();

                ImportDelegate importCodeAction = new ImportDelegate(ImportCode);

                scriptEngine.AddObject("Import", importCodeAction);

                scriptEngine.AddCode(addCode);

                script.AddCode();

                //MessageBox.Show(scriptEngine.Eval("Add 1, 2").ToString());
            }
            catch (Exception ex)
            {

            }
        }

        [ComVisible(true)]
        public void ImportCode(string filePath)
        {

        }

        [ComVisible(true)]
        public delegate void ImportDelegate(string filePath);

    }

    [ComVisible(true)]
    public class HostObject
    {
        public void Echo(string text)
        {
            
        }
    }
}
