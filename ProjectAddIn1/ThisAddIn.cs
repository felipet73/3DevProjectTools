using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;


using Project = Microsoft.Office.Interop.MSProject;

using Microsoft.Office.Tools;
using Microsoft.Office.Interop.MSProject;
using System.Windows.Forms;

namespace ProjectAddIn1
{
    public partial class ThisAddIn
    {
        private UserControl1 myUserControl1;
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public string anterior;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Application.WindowSelectionChange += new EventHandler(Application_WindowSelectionChange);
            //Application.WindowSelectionChange += new Microsoft.Office.Interop.MSProject._EP(Application_WindowSelectionChange);
            Application.WindowSelectionChange += new Microsoft.Office.Interop.MSProject._EProjectApp2_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        private void Application_WindowSelectionChange(Window Window, Selection sel, object selType)
        {
            //throw new NotImplementedException();
            try
            {
                // MessageBox.Show(Globals.ThisAddIn.Application.ActiveCell.Task.Name);


                if (Globals.ThisAddIn.Application.ActiveCell.Task.Text1 != "" && Globals.ThisAddIn.Application.ActiveCell.Task.Text1!=anterior) {
                    myUserControl1.ejecutar(Globals.ThisAddIn.Application.ActiveCell.Task.Text1);
                    anterior = Globals.ThisAddIn.Application.ActiveCell.Task.Text1;
                }
            }
            catch { 
            
            }

        }

        public void activar()
        {
            myUserControl1 = new UserControl1();


            Microsoft.Office.Tools.CustomTaskPaneCollection customPaneCollection;
            customPaneCollection = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "Panes", "Panes", this);

            Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane = customPaneCollection.Add(myUserControl1, "Title");
            myCustomTaskPane.Visible = true;

            /*myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1,
                "New Task Pane");*/

            myCustomTaskPane.DockPosition =
                Office.MsoCTPDockPosition.msoCTPDockPositionFloating;
            myCustomTaskPane.Height = 500;
            myCustomTaskPane.Width = 500;

            myCustomTaskPane.DockPosition =
                Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            myCustomTaskPane.Width = 300;

            myCustomTaskPane.Visible = true;
            myCustomTaskPane.DockPositionChanged +=
                new EventHandler(myCustomTaskPane_DockPositionChanged);

        }

        private void myCustomTaskPane_DockPositionChanged(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }



        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
