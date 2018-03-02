using BaseProject.Report;
using System;
using System.Collections.Generic;

namespace BaseProject
{
    public class Should
    {
        public static void Equal(Iteration reporter, String actual, String expected)
        {
            reporter.Chapter.Step.Actions.Add(new Act(String.Format("Verify '{0}' Equals '{1}'", expected, actual)));

            if (!String.Equals(actual, expected))
            {
                throw new Exception(String.Format("Not Equal {0} : {1}", actual, expected));
            }
        }

        public static void Equal(Iteration reporter, DateTime actual, DateTime expected, String name = "")
        {
            reporter.Chapter.Step.Actions.Add(new Act(String.Format("Verify '{0}' Equals '{1}'", expected, actual)));

            if (!String.Equals(actual, expected))
            {
                throw new Exception(String.Format("Not Equal {0} : {1}", actual, expected));
            }
        }

        public static void NullOrEmpty(Iteration reporter, String data)
        {
            reporter.Chapter.Step.Actions.Add(new Act(String.Format("Verify Null or Empty '{0}'", data)));

            if (!String.IsNullOrEmpty(data) || !String.IsNullOrWhiteSpace(data))
            {
                throw new Exception(String.Format("Data is not Null or Empty"));
            }
        }

        public static void Equal(Iteration reporter, Int64 first, Int64 second)
        {
            if (!(first == second))
            {
                throw new Exception(String.Format("Not Equal {0} : {1}", first, second));
            }
        }

        public static void Equal(Iteration reporter, bool first, bool second)
        {
            reporter.Chapter.Step.Actions.Add(new Act(String.Format("Verify '{0}' Equals '{1}'", first, second)));

            if (!(first == second))
            {
                throw new Exception(String.Format("Not Equal {0} : {1}", first, second));
            }
        }

        public static void Equal(Iteration reporter, Decimal first, Decimal second)
        {
            if (!first.Equals(second))
            {
                throw new Exception(String.Format("Not Equal {0} : {1}", first, second));
            }
        }

        public class Be
        {
            public static void AlphabeticalOrder(IList<String> items)
            {
                String lastItem = String.Empty;
                foreach (String item in items)
                {
                    if (lastItem == String.Empty)
                    {
                        lastItem = item;
                    }

                    if (lastItem.CompareTo(item) > 0)
                    {
                        throw new Exception(String.Format("Item {0} not in alphabetical order", item));
                    }
                }
            }
        }

        public class Not
        {
            public static void NullOrEmpty(String data)
            {
                if (String.IsNullOrEmpty(data) || String.IsNullOrWhiteSpace(data))
                {
                    throw new Exception(String.Format("Data is Null or Empty"));
                }
            }

            public static void Equal(Iteration reporter, String first, String second)
            {
                reporter.Chapter.Step.Actions.Add(new Act(String.Format("Verify '{0}' Not Equals '{1}'", first, second)));

                if (String.Equals(first, second))
                {
                    throw new Exception(String.Format("Equal {0} : {1}", first, second));
                }
            }

            public static void Equal(Iteration reporter, Int64 first, Int64 second)
            {
                reporter.Chapter.Step.Actions.Add(new Act(String.Format("Verify '{0}' Not Equals '{1}'", first, second)));

                if (first == second)
                {
                    throw new Exception(String.Format("Not Equal {0} : {1}", first, second));
                }
            }
        }
    }
}
