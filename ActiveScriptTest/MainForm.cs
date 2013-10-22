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
    using System.Diagnostics;

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

            string codeWithError =
                "Dim a : a = 1 / 0";

            ActiveScriptEngine scriptEngine = new ActiveScriptEngine(VBScript.ProgID);

            scriptEngine.ScriptErrorOccurred += new ScriptErrorOccurredDelegate(scriptEngine_ScriptErrorOccurred);

            scriptEngine.AddCode(addFunc, "Math");
            scriptEngine.AddCode(subtract, "Math");

            scriptEngine.AddCode(
                "Public Function Add(a, b) " + Environment.NewLine +
                "   Add = Math.Add(a, b) " + Environment.NewLine +
                "End Function");

            scriptEngine.AddCode(echo);

            scriptEngine.Initialize();

            scriptEngine.Start();

            dynamic script = scriptEngine.GetScriptHandle();

            Log(script.Math.Subtract(script.Add(1, 3), 2));

            scriptEngine.AddObject("WScript", new HostObject());

            script.SayHello();

            try
            {
                scriptEngine.AddCode(codeWithError);
            }
            catch
            {
                ScriptErrorInfo error = scriptEngine.LastError;

                Log("Catch block");
                Log(error.Description);
            }
            finally
            {
                scriptEngine.Dispose();
            }
        }

        private void scriptEngine_ScriptErrorOccurred(ActiveScriptEngine sender, ScriptErrorInfo error)
        {
            Log("event block");
            Log(error.Description);
        }

        public static void Log(object text)
        {
            Debug.WriteLine(text);
        }
    }

    [ComVisible(true)]
    public class HostObject
    {
        public void Echo(string text)
        {
            MainForm.Log(text);
        }
    }
}
