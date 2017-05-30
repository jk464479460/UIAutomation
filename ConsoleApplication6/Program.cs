using System;
using System.Windows.Automation;

namespace ConsoleApplication6
{
    class CalcAutomationClient
    {
        static AutomationElement calcWindow = null;
        static string resultTextAutoID = "150";
        static string btn5AutoID = "135";
        static string btn3AutoID = "133";
        static string btn2AutoID = "132";
        static string btnPlusAutoID = "93";
        static string btnSubAutoID = "94";
        static string btnEqualAutoID = "121";

        void OnWindowOpenOrClose(object src, AutomationEventArgs e)
        {

            if (e.EventId != WindowPattern.WindowOpenedEvent)
            {
                return;
            }

            AutomationElement sourceElement;
            try
            {
                sourceElement = src as AutomationElement;
                //Check the event source is caculator or not.
                //In production code, string should be read from resource to support localization testing.
                if (sourceElement.Current.Name == "计算器")
                {
                    calcWindow = sourceElement;
                }

            }
            catch (ElementNotAvailableException)
            {
                return;
            }
            //Start testing
            ExecuteTest();

        }

        void ExecuteTest()
        {
            //Execute 3+5-2"
            //Invoke ExecuteButtonInvoke function to click buttons
            ExecuteButtonInvoke(btn3AutoID);
            ExecuteButtonInvoke(btnPlusAutoID);
            ExecuteButtonInvoke(btn5AutoID);
            ExecuteButtonInvoke(btnSubAutoID);
            ExecuteButtonInvoke(btn2AutoID);
            ExecuteButtonInvoke(btnEqualAutoID);
            //Invoke GetCurrentResult function to read caculator output
            if (GetCurrentResult() == "6")
            {
                Console.WriteLine("Execute Pass!");
                return;
            }

            Console.WriteLine("Execute Fail!");

        }
        void ExecuteButtonInvoke(string automationID)
        {
            //Create query condition object, there are two conditions.
            //1. Check AutomationID
            //2. Check Control Type
            Condition conditions = new AndCondition(
                new PropertyCondition(AutomationElement.AutomationIdProperty, automationID),
                 new PropertyCondition(AutomationElement.ControlTypeProperty,
                                 ControlType.Button));

            AutomationElement btn = calcWindow.FindAll(TreeScope.Descendants, conditions)[0];
            //Obtain the InvokePattern interface
            InvokePattern invokeptn = (InvokePattern)btn.GetCurrentPattern(InvokePattern.Pattern);
            //Click button by Invoke interface
            invokeptn.Invoke();
        }

        string GetCurrentResult()
        {
            Condition conditions = new AndCondition(
                new PropertyCondition(AutomationElement.AutomationIdProperty, resultTextAutoID),
                 new PropertyCondition(AutomationElement.ControlTypeProperty,
                                 ControlType.Text));
            AutomationElement btn = calcWindow.FindAll(TreeScope.Descendants, conditions)[0];
            //Read name property of Text control. The name property is the output.
            return btn.Current.Name;
        }

        static void Main(string[] args)
        {
            CalcAutomationClient autoClient = new CalcAutomationClient();

            //Create callback for new Window open event. Test should run only when the main Window shows.

            AutomationEventHandler eventHandler = new AutomationEventHandler(autoClient.OnWindowOpenOrClose);

            //Attach the event with desktop element and start listening.

            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, TreeScope.Children, eventHandler);

            //Start caculator. When new window opens, the new window open event should fire.

            System.Diagnostics.Process.Start("calc.exe");

            //Wait execution

            Console.ReadLine();
        }
    }
}
