using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using TechTalk.SpecFlow;


namespace WPFtelerik.Steps
{
    [Binding]
    public sealed class StepDefinition1

    {
        AutomationElement aWPFapp = null;
        AutomationElement aeCustom1 = null;
        AutomationElementCollection aeAllControl;
        string fileName = @"C:\temp\Griddata.txt";

        [Given(@"I launch the application")]
        public void GivenILaunchTheApplication()
        {
            try
            {
                Process p = null;
                p=Process.Start(new ProcessStartInfo(@"C:/Users/hello/Desktop/Telerik UI for WPF - Demos.appref-ms") { UseShellExecute = true });
                //p = Process.Start("C:/Users/hello/Desktop/Telerik UI for WPF - Demos.appref-ms");

                int ct = 0;
                do
                {
                    ++ct;
                    Thread.Sleep(100);
                } while (p == null && ct < 50);
                if (p == null)
                    throw new Exception("Failed to find WPF process");
             
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal: " + ex.Message);
            }
        }

        [Given(@"I click on GridView Button")]
        public void GivenIClickOnGridViewButton()
        {
            try {
               
                AutomationElement aeDesktop = null;
                
                aeDesktop = AutomationElement.RootElement;
            if (aeDesktop == null)
                throw new Exception("Unable to get Desktop");
            
            int numWaits = 0;
            do
            {
                aWPFapp = aeDesktop.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "WPF Controls Examples - GridView, ChartView, ScheduleView, RichTextBox, Map, Code Samples"));
                ++numWaits;
                Thread.Sleep(200);
            }
            while (aWPFapp == null && numWaits < 50);
            if (aWPFapp == null)
                throw new Exception("Failed to find WPF main window");
            aeCustom1 = aWPFapp.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "rootHome"));
            if (aeCustom1 == null)
                throw new Exception("No custom1 found");
            

            AutomationElement aeCustom2 = aeCustom1.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "rootHamburger"));
            if (aeCustom2 == null)
                throw new Exception("No custom2 found");
            
            AutomationElement aeHamview = aeCustom2.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "HamburgerExpander"));
            if (aeHamview == null)
                throw new Exception("No hamview found");
            
            
            aeAllControl = aeHamview.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "rootHighlightedControls"));
            if (aeAllControl == null)
                throw new Exception("No Root Control found");
            
            AutomationElement aeReqControl = null;
            aeReqControl = aeAllControl[1];

            AutomationElementCollection aeAllText = null;
            aeAllText = aeReqControl.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.FrameworkIdProperty, "WPF"));
            if (aeAllText == null)
                throw new Exception("No Button collection");
            

            AutomationElement gridButton = aeAllText[2];
            if (gridButton == null)
                throw new Exception("Cannot find the button");
            else
            {
                InvokePattern ipClickButton1 = (InvokePattern)gridButton.GetCurrentPattern(InvokePattern.Pattern);
                ipClickButton1.Invoke();
            }

            Thread.Sleep(1000);
        }
        catch(Exception ex)
            {
                Console.WriteLine("Fatal: " + ex.Message);
            }
        }

        [When(@"the GridView Screen Comes")]
        public void WhenTheGridViewScreenComes()
        {
            try 
            {
                var aedatagrid = aWPFapp.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid));
                if (aedatagrid == null)
                    throw new Exception("Cannot find the grid details GRID");
            }
            catch(Exception ex)
            {
                Console.WriteLine("Fatal: " + ex.Message);
            }
        }

        [Then(@"I get all the data from Grid and store it in Text file")]
        public void ThenIGetAllTheDataFromGridAndStoreItInTextFile()
        {
            try
            {
               
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
                using StreamWriter sw = File.CreateText(fileName);
                {
                    sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                    sw.WriteLine("Author: Chandni Patra");
                    sw.WriteLine("File Created ");
                    var aedatagrid = aWPFapp.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataGrid));
                    if (aedatagrid == null)
                        throw new Exception("Cannot find the grid details GRID");
                    else
                        sw.WriteLine("grid details found");
                    var rows = aedatagrid.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem));
                    foreach (AutomationElement row in rows)
                    {
                        var findRow = row.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
                        sw.WriteLine("===============");
                        foreach (AutomationElement a in findRow)
                        {
                            sw.WriteLine(a.Current.Name);
                        }

                    }
                    sw.WriteLine("===============");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal: " + ex.Message);
            }

        }

    }
}
