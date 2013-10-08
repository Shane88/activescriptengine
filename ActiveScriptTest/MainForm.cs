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

            try
            {
                ActiveScriptEngine scriptEngine = new ActiveScriptEngine("VBScript");

                scriptEngine.AddCode(addFunc, "Math");
                scriptEngine.AddCode(addFunc, "Math");
                scriptEngine.AddCode(subtract, "Math");

                scriptEngine.AddCode(
                    "Public Function Add(a, b) " + Environment.NewLine +
                    "   Add = Math.Add(a, b) " + Environment.NewLine +
                    "End Function");


                scriptEngine.Start();

                dynamic script = scriptEngine.GetIDispatch();

                MessageBox.Show(script.Math.Subtract(script.Add(1, 3), 2).ToString());
            }
            catch (Exception ex)
            {

            }
        }
    }
}
