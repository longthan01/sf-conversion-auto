using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace SF_ConversionAuto
{
    public abstract class ControlValueController
    {
        public string Value { get; set; }
        public string ControlId { get; set; }
        public string ProcessName
        {
            get;
        }

        protected Process Process;
        protected AutomationElement MainWindow;
        protected ControlValueController ChainedHandler;
        public ControlValueController(string value, string controlId, string processName)
        {
            Value = value;
            ControlId = controlId;
            ProcessName = processName;
            var procs = Process.GetProcessesByName(processName);
            if (procs.Length == 0)
            {
                throw new Exception($"Could not find process with name {processName}");
            }

            Process = procs[0];
            Condition yourCondition = new PropertyCondition(AutomationElement.ProcessIdProperty, Process.Id);
            MainWindow =
                AutomationElement.RootElement.FindFirst(System.Windows.Automation.TreeScope.Children,
                    yourCondition);
        }

        public void SetValue()
        {
            this.Execute();
            if (this.ChainedHandler != null)
            {
                this.ChainedHandler.Execute();
            }
        }

        protected abstract void Execute();

        public ControlValueController SetChain(ControlValueController handler)
        {
            this.ChainedHandler = handler;
            return this;
        }
    }

    public class TextBoxControlValueController : ControlValueController
    {
        public TextBoxControlValueController(string value, string controlId, string processName) : base(value, controlId, processName)
        {
        }

        protected override void Execute()
        {
            Condition controlCondition = new PropertyCondition(AutomationElement.AutomationIdProperty, ControlId);
            var element = MainWindow.FindFirst(TreeScope.Children, controlCondition);
            if (element != null)
            {
                try
                {
                    object pattern = element.GetCurrentPattern(ValuePattern.Pattern);
                    if (pattern != null && pattern is ValuePattern)
                    {
                        ValuePattern valuePtr = pattern as ValuePattern;
                        valuePtr.SetValue(Value);
                    }
                }
                catch (Exception)
                {

                }
            }
            ChainedHandler.SetValue();
        }
    }

    public class ComboboxControlValueController : ControlValueController
    {
        public ComboboxControlValueController(string value, string controlId, string processName) : base(value, controlId, processName)
        {
        }

        protected override void Execute()
        {
            Condition controlCondition = new PropertyCondition(AutomationElement.AutomationIdProperty, ControlId);
            var element = MainWindow.FindFirst(TreeScope.Descendants, controlCondition);
            if (element != null)
            {
                ExpandCollapsePattern expandPattern = null;
                try
                {
                    expandPattern = (ExpandCollapsePattern)element.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                    expandPattern.Expand();
                    //expandPattern.Collapse();
                }
                catch (Exception)
                {

                }
                AutomationElementCollection comboboxItems = element.FindAll(TreeScope.Descendants, 
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));

                foreach (AutomationElement automationElement in comboboxItems)
                {
                    string elementValue = automationElement.Current.Name;
                    if(!string.IsNullOrEmpty(elementValue) && elementValue.Trim().ToLower().Equals(this.Value.Trim().ToLower()))
                    {
                        //Finding the pattern which need to select
                        SelectionItemPattern selectPattern = (SelectionItemPattern)automationElement.GetCurrentPattern(SelectionItemPattern.Pattern);
                        selectPattern.Select();
                    }
                }
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var parameters = ParseParams(args);
                var processName = parameters.FirstOrDefault(x => x.Key == "-procName").Value;
                var controlId = parameters.FirstOrDefault(x => x.Key == "-controlId").Value;
                var controlValue = parameters.FirstOrDefault(x => x.Key == "-controlValue").Value;
                var controller = new TextBoxControlValueController(controlValue, controlId, processName)
                    .SetChain(new ComboboxControlValueController(controlValue, controlId, processName));
                controller.SetValue();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Out.WriteLine($"ERROR: {ex}");
                Console.ForegroundColor = ConsoleColor.White;
            }
        }
        private static Dictionary<string, string> ParseParams(string[] args)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            for (int i = 0; i < args.Length; i++)
            {
                if (i + 1 < args.Length)
                {
                    parameters.Add(args[i], args[i + 1]);
                    i++;
                }
                else
                {
                    parameters.Add(args[i], "");
                }
            }
            return parameters;
        }
    }
}
