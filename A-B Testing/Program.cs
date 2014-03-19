using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using MathNet.Numerics.Statistics;
using MathNet.Numerics.Distributions;

namespace ChallengePart1
{
    class Program
    {
        public static String inputFile = @"C:\Optmzr\ADPERFORMANCEREPORT.CSV";
        public static String OutputFile = @"C:\Optmzr\OUTPUT.CSV";
        public static String SheetName1 = @"ADPERFORMANCEREPORT";
        public static String SheetName2 = @"OUTPUT";
        static void Main(string[] args)
        {
            try
            {
                Application oXL = new Application();
                oXL.DisplayAlerts = false;
#if DEBUG
                oXL.Visible = true;
#else
	oXL.Visible = false; 
#endif
                //Open the Excel File

                Workbook oWB = oXL.Workbooks.Open(inputFile);

                _Worksheet oSheet = oWB.Sheets[SheetName1];

                oSheet = oWB.Sheets[1]; // Gives the reference to the first sheet

                oSheet = oWB.ActiveSheet; // Gives the reference to the current opened sheet

                oSheet.Activate();

                oSheet.Copy(Type.Missing, Type.Missing);
                _Worksheet oSheet1 = oXL.Workbooks[2].Sheets[1];
                oSheet1.Name = SheetName2;

                try
                {
                    //We already know the Address Range of the cells

                    String start_range = "A2";
                    String end_range = "A164";

                    Object[,] valuesA = oSheet1.get_Range(start_range, end_range).Value2;

                    start_range = "B2";
                    end_range = "B164";

                    Object[,] valuesB = oSheet.get_Range(start_range, end_range).Value2;


                    start_range = "C2";
                    end_range = "C164";

                    Object[,] valuesC = oSheet.get_Range(start_range, end_range).Value2;


                    start_range = "D2";
                    end_range = "D164";

                    Object[,] valuesD = oSheet.get_Range(start_range, end_range).Value2;

                    //int t = valuesB.GetLength(0);

                    object duplicate = null;
                    int index = 1;
                    int count = 0;
                    int i = 1;
                    for (i = 1; i <= valuesA.GetLength(0); i++)
                    {
                        if (duplicate == null)
                        {
                            duplicate = valuesA[i, 1];
                            index = i;
                        }
                        if (Convert.ToInt64(valuesA[i, 1]) == Convert.ToInt64(duplicate))
                        {
                            count++;
                            duplicate = valuesA[i, 1];
                            continue;
                        }
                        double[] probSuc = new double[count];
                        //double max = Double.MinValue;
                        int r = 1;
                        int[] successes = new int[count];
                        int[] failures = new int[count];
                        int[] trials = new int[count];
                        for (int k = index; k <= (index + count - 1); k++)
                        {
                            double s = Convert.ToInt32(valuesC[k, 1]);
                            double n = Convert.ToInt32(valuesD[k, 1]);
                            double f = Convert.ToInt32(valuesD[k, 1]) - s;
                            trials[r - 1] = Convert.ToInt32(n);
                            successes[r - 1] = Convert.ToInt32(s);
                            failures[r - 1] = Convert.ToInt32(f);
                            probSuc[r - 1] = s / n;
                            r++;
                        }                      
                        
                        r = 1;
                        for (int k = index; k <= (index + count - 1); k++)
                        {
                            var samples = new Binomial(probSuc[r-1], trials[r-1]);
                            var samples1 = new Bernoulli(probSuc[r - 1]);                            
                            double prob = samples.Probability(0);

                            if (prob < 0.05)
                            {
                                oSheet1.Cells[k + 1, 5] = "WINNER";
                            }
                            else
                            {
                                oSheet1.Cells[k + 1, 5] = "LOSSER";
                            }
                            r++;
                        }
                        i--;
                        duplicate = null;
                        count = 0;
                    }
                    if (count > 0 && index <= valuesA.GetLength(0))
                    {
                        double[] probSuc = new double[count];
                        //double max = Double.MinValue;
                        int r = 1;
                        int[] successes = new int[count];
                        int[] failures = new int[count];
                        int[] trials = new int[count];
                        for (int k = index; k <= (index + count - 1); k++)
                        {
                            double s = Convert.ToInt32(valuesC[k, 1]);
                            double n = Convert.ToInt32(valuesD[k, 1]);
                            double f = Convert.ToInt32(valuesD[k, 1]) - s;
                            trials[r - 1] = Convert.ToInt32(n);
                            successes[r - 1] = Convert.ToInt32(s);
                            failures[r - 1] = Convert.ToInt32(f);
                            probSuc[r - 1] = s / n;
                            r++;
                        }
                        r = 1;
                        for (int k = index; k <= (index + count - 1); k++)
                        {
                            var samples = new Binomial(probSuc[r - 1], trials[r - 1]);
                            double prob = samples.Probability(0);

                            if (prob < 0.05)
                            {
                                oSheet1.Cells[k + 1, 5] = "WINNER";
                            }
                            else
                            {
                                oSheet1.Cells[k + 1, 5] = "LOSSER";
                            }
                        }
                    }

                }
                catch (Exception e)
                {
                    var errors = e.Message;
                }
                finally
                {
                    oSheet1.SaveAs(OutputFile);
                    oXL.Quit();

                    Marshal.ReleaseComObject(oSheet);
                    Marshal.ReleaseComObject(oSheet1);
                    Marshal.ReleaseComObject(oWB);
                    Marshal.ReleaseComObject(oXL);

                    oSheet = null;
                    oSheet1 = null;
                    oWB = null;
                    oXL = null;
                    GC.GetTotalMemory(false);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.GetTotalMemory(true);
                }
            }
            catch (Exception e)
            {
                var errors = e.Message;
            }
        }
    }
}
