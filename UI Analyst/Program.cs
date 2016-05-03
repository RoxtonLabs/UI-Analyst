using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Automation;    //Requires reference made to C:\Program Files\Reference Assemblies\Microsoft\Framework\v3.0\UIAutomationClient.dll and UIAutomationTypes.dll.
//using System.Windows.Controls;  //Requires reference to Assemblies>Framework>PresentationFramework.

namespace UI_Analyst
{
    class Program
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);


        static void Main(string[] args)
        {
            Console.WriteLine("Gathering window titles...\n");
            List<IntPtr> windowHandles = WindowHandles();
            if (windowHandles.Count == 0)
            {   //This should never get called - the app itself is a window!
                Console.WriteLine("No windows found.");
            }
            else
            {
                Console.WriteLine("{0} window{1} found:\n", windowHandles.Count, (windowHandles.Count > 1) ? "s" : "");
                for (int i = 0; i < windowHandles.Count; i++)
                {
                    string windowTitle = WindowTitle(windowHandles[i]);
                    if (windowTitle != null)
                    {
                        Console.WriteLine("{0}:\t{1}",i+1, windowTitle);
                    }
                }
                Console.WriteLine("\nChoose a window to analyse ({0}-{1}):", "1", windowHandles.Count);

                ChooseWindow:
                string input = Console.ReadLine();
                int numInput = 0;
                if (!Int32.TryParse(input, out numInput) || (numInput < 1 || numInput > windowHandles.Count))
                {
                    Console.WriteLine("Please choose a window between {0} and {1}:", "1", windowHandles.Count);
                    goto ChooseWindow;
                }

                //Now create a list of all the UI elements in the chosen window
                AutomationElement windowElement = AutomationElement.FromHandle(windowHandles[numInput - 1]);
                AutomationElementCollection windowElements = windowElement.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
                foreach (AutomationElement ae in windowElements)
                {
                    Console.WriteLine("ClassName: {0}", ae.Current.ClassName);
                    Console.WriteLine("Text: {0}", AutomationExtensions.GetText(ae));
                    Console.WriteLine("");
                }
                //TODO loop down through each element, increasing the number of \t each time, and thus print the entire tree
                //TODO print that tree to a file
            }

            Console.WriteLine("\n\nPress <ENTER> to exit...");
            Console.ReadLine();
        }

        /// <summary>
        /// Attempts to return the title of the window with the specified handle.
        /// If no window is found with that handle, returns null.
        /// </summary>
        /// <param name="hWnd">The handle of the window to find.</param>
        /// <returns>The title of the window, or null if no window could be found that matches the supplied handle.</returns>
        static string WindowTitle(IntPtr hWnd)
        {
            int capacity = GetWindowTextLength(hWnd) * 2;
            StringBuilder stringBuilder = new StringBuilder(capacity);
            GetWindowText(hWnd, stringBuilder, stringBuilder.Capacity);
            if (stringBuilder.Length > 0)
            {
                return stringBuilder.ToString();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Returns a list of all the handles of processes that have main windows.
        /// </summary>
        static List<IntPtr> WindowHandles()
        {
            List<IntPtr> toReturn = new List<IntPtr>();
            Process[] processList = Process.GetProcesses();
            foreach (Process p in processList)
            {
                if (!String.IsNullOrEmpty(p.MainWindowTitle))
                {
                    toReturn.Add(p.MainWindowHandle);
                }
            }
            return toReturn;
        }
    }

    public static class AutomationExtensions
    {

        public static string GetText(this AutomationElement element)
        {
            object patternObj;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r'); // often there is an extra '\r' hanging off the end.
            }
            else
            {
                return element.Current.Name;
            }
        }

        /// <summary>
        /// Attempts to expand the given AutomationElement. Returns null if successful or an error message if not.
        /// </summary>
        /// <param name="element">The AutomationElement to expand.</param>
        /// <returns>Null if sucecssful or an error message if not.</returns>
        public static string ExpandElement(AutomationElement element)
        {
            string toReturn = string.Format("Unable to expand element {0}: reason unknown.", GetText(element));
            try
            {
                ExpandCollapsePattern ecp = null;
                ecp = element.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;

                if (ecp.Current.ExpandCollapseState == ExpandCollapseState.Collapsed ||
                    ecp.Current.ExpandCollapseState == ExpandCollapseState.PartiallyExpanded)
                {
                    ecp.Expand();
                    toReturn = null;
                }
            }
            catch (Exception ex)
            {
                toReturn = string.Format("Unable to expand element {0}: {1}.", GetText(element), ex.Message);
            }
            return toReturn;
        }
    }

}
