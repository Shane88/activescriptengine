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
            try
            {
                ActiveScriptEngine scriptEngine = new ActiveScriptEngine("VBScript");

                scriptEngine.AddCode(
                    "Public Function Add(a, b) " + Environment.NewLine +
                    "   Add = a + b " + Environment.NewLine +
                    "End Function");

                scriptEngine.Run();

                dynamic script = scriptEngine.GetIDispatch();

                MessageBox.Show(script.Add(1, 5).ToString());
            }
            catch (Exception ex)
            {

            }
        }
    }
}
