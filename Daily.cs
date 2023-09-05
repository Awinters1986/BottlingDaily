class Daily
    {
        public static List<string> dtCategories;
        public static Boolean Monday = false;
        public static Boolean twoTanks = false;
        public static Boolean threeTanks = false;

        public static void Delete(string FolderName)
        {
            DirectoryInfo dir = new DirectoryInfo(FolderName);

            foreach (FileInfo file in dir.GetFiles())
            {
                file.Delete();
            }

            foreach (DirectoryInfo di in dir.GetDirectories())
            {
                Delete(di.FullName);
                di.Delete();
            }
        }
        public static void Copy(string sourceDirectory, string targetDirectory)
        {
            DirectoryInfo dirSource = new DirectoryInfo(sourceDirectory);
            DirectoryInfo dirTarget = new DirectoryInfo(targetDirectory);

            CopyToday(dirSource, dirTarget);
        }

        public static void CopyToday(DirectoryInfo source, DirectoryInfo target)
        {
            FileInfo[] files = source.GetFiles("*", SearchOption.AllDirectories);

            foreach (FileInfo file in files)
            {
                //Copy file if the modified date matches either today or yesterday
                if ((file.LastWriteTime.Date == DateTime.Today.Date) | file.LastWriteTime.Date == DateTime.Today.AddDays(-1))
                {
                    Console.WriteLine(@"Copying {0}\{1}", target.FullName, file.Name);
                    file.CopyTo(Path.Combine(target.FullName, file.Name), true);
                }
            }
        }
        public static void createReportDirectory()
        {
            string reportDirectory = @"C:\AWinters\Bottling Daily\Today\Reports";
            Directory.CreateDirectory(reportDirectory);
        }

        public static void writeProductionCSV(List<ProductionLog> workOrders)
        {
            using (var textWriter = new StreamWriter(@"C:\AWinters\Bottling Daily\Today\Reports" + "\\production.csv"))
            {
                var writer = new CsvWriter(textWriter, CultureInfo.InvariantCulture);
            }
        }

        public static void writeOEECSV(List<ProductionLog> workOrders)
        {
            using (var textWriter = new StreamWriter(@"C:\AWinters\Bottling Daily\Today\Reports" + "\\oee.csv"))
            {
                var writer = new CsvWriter(textWriter, CultureInfo.InvariantCulture);
            }
        }

        public static void writeDowntimeCSV(string[,] dtValues)
        {
            using (var textWriter = new StreamWriter(@"C:\AWinters\Bottling Daily\Today\Reports" + "\\downtime.csv"))
            {
                var writer = new CsvWriter(textWriter, CultureInfo.InvariantCulture);
                for (int i = 0; i < dtValues.Length; i++)
                {
                    for (int j = 0; j < 7; j++)
                    {
                        writer.WriteField(dtValues[i, j]);
                    }
                    writer.NextRecord();
                }
            }
        }

        public static void writeConsumablesCSV()
        {
            using (var textWriter = new StreamWriter(@"C:\AWinters\Bottling Daily\Today\Reports" + "\\consumables.csv"))
            {
                var writer = new CsvWriter(textWriter, CultureInfo.InvariantCulture);
            }
        }
        public static void loadDowntimeCategoriesList()
        {
            dtCategories = new List<string>();
            dtCategories.Add("Fill Height/Cap Detector");
            dtCategories.Add("FCS");
            dtCategories.Add("Proc Tk/Chg");
            dtCategories.Add("Manning/Supplies");
            dtCategories.Add("Foam");
            dtCategories.Add("Proof Cks/Lab");
            dtCategories.Add("QA External");
            dtCategories.Add("QA Internal");
            dtCategories.Add("Pal/Depal");
            dtCategories.Add("Unloader");
            dtCategories.Add("Bulk Unloader");
            dtCategories.Add("Cleaner");
            dtCategories.Add("Filler");
            dtCategories.Add("Capper");
            dtCategories.Add("Checkweighter/Taper");
            dtCategories.Add("Team Performance");
            dtCategories.Add("Bottle Orientor");
            dtCategories.Add("Labeler");
            dtCategories.Add("Wash");
            dtCategories.Add("Case Packer");
            dtCategories.Add("Partitioner");
            dtCategories.Add("Case Erector");
            dtCategories.Add("Polypak");
            dtCategories.Add("Case Line");
            dtCategories.Add("State Code");
            dtCategories.Add("Case Printer");
            dtCategories.Add("Laser Dater");
            dtCategories.Add("Other");
        }

        public static void Main(string[] args)
        {
            string sourceDirectory = @"C:\Bottling Downtime Sheets";
            string targetDirectory = @"C:\AWinters\Bottling Daily\Today";

            //TODO: Add Monday logic
            loadDowntimeCategoriesList(); //Load all downtime categories into a list. This will allow ease of pulling this category after loading downtime information from excel.
            Delete(targetDirectory); //Delete yesterday's files from Personal folder to lessen the interference with today's files
            Copy(sourceDirectory, targetDirectory); //copy files from Bottling Downtime to Centralized Directory
            MergeSheets(targetDirectory);
        }

        public static void MergeSheets(string dirPath)
        {
            DirectoryInfo directory = new DirectoryInfo(dirPath);
            FileInfo[] files = directory.GetFiles("*", SearchOption.AllDirectories);
            createReportDirectory();
            List<ProductionLog> workOrders = new List<ProductionLog>();

            try
            {
                Excel.Application app = new Excel.Application();
                app.Visible = true;

                foreach (FileInfo file in files)
                {
                    Boolean setupWO = true;
                    string fileName = dirPath + "\\" + file.Name;
                    Excel._Workbook wb = app.Workbooks.Open(fileName);
                    Console.WriteLine(fileName);

                    //~~Production Log Read~~//
                    Excel._Worksheet prodws = wb.Sheets[1];
                    Excel.Range prodrng = prodws.UsedRange;
                    try
                    {
                        var prodDate = (string)(prodws.Cells[2, 12] as Excel.Range).Text;
                        if (DateTime.Parse(prodDate) == DateTime.Today.AddDays(-1))
                        {
                            var workOrder = (string)(prodws.Cells[6, 1] as Excel.Range).Text;
                            var item = (string)(prodws.Cells[6, 4] as Excel.Range).Text;
                            var description = (string)(prodws.Cells[6, 7] as Excel.Range).Text;
                            var sizeCase = (string)(prodws.Cells[6, 11] as Excel.Range).Text;
                            var shift = (string)(prodws.Cells[6, 13] as Excel.Range).Text;
                            var line = (string)(prodws.Cells[6, 14] as Excel.Range).Text;
                            var alcoholNum = (string)(prodws.Cells[8, 2] as Excel.Range).Text;

                            ProductionLog wo = new ProductionLog(prodDate, workOrder, item, sizeCase, shift, line, alcoholNum);
                            workOrders.Add(wo);

                            //~~Tank Information~~//
                            if ((prodws.Cells[8, 6] as Excel.Range).Text != null)
                            {
                                setupWO = false;
                                wo.Tank = (string)(prodws.Cells[8, 6] as Excel.Range).Text;
                                wo.Cases = (string)(prodws.Cells[8, 11] as Excel.Range).Text;
                                wo.Reduction = (string)(prodws.Cells[10, 3] as Excel.Range).Text;
                                wo.Serial = (string)(prodws.Cells[10, 7] as Excel.Range).Text;
                                wo.Gallons = (string)(prodws.Cells[10, 11] as Excel.Range).Text;
                            }
                            //If two tanks are used
                            if (((prodws.Cells[14, 6] as Excel.Range).Text != null) && (setupWO = false))
                            {
                                twoTanks = true;
                                wo.Tank2 = (string)(prodws.Cells[14, 6] as Excel.Range).Text;
                                wo.Cases2 = (string)(prodws.Cells[14, 11] as Excel.Range).Text;
                                wo.Reduction2 = (string)(prodws.Cells[16, 3] as Excel.Range).Text;
                                wo.Serial2 = (string)(prodws.Cells[16, 7] as Excel.Range).Text;
                                wo.Gallons2 = (string)(prodws.Cells[16, 11] as Excel.Range).Text;
                            }
//If three tanks are used
                            if (((prodws.Cells[20, 6] as Excel.Range).Value != null) && (setupWO = false))
                            {
                                threeTanks = true;
                                wo.Tank3 = (string)(prodws.Cells[20, 6] as Excel.Range).Text;
                                wo.Cases3 = (string)(prodws.Cells[20, 11] as Excel.Range).Text;
                                wo.Reduction3 = (string)(prodws.Cells[22, 3] as Excel.Range).Text;
                                wo.Serial3 = (string)(prodws.Cells[22, 7] as Excel.Range).Text;
                                wo.Gallons3 = (string)(prodws.Cells[22, 11] as Excel.Range).Text;
                            }

                            wo.TotalCases = (string)(prodws.Cells[27, 13] as Excel.Range).Text;

                            //~~Manning Information~~//
                            if (((prodws.Cells[27, 1] as Excel.Range).Value != null)) //Checking if setup minutes are included
                            {
                                wo.SetupStart = (string)(prodws.Cells[27, 1] as Excel.Range).Text;
                                wo.SetupEnd = (string)(prodws.Cells[27, 2] as Excel.Range).Text;
                                wo.SetupMin = (string)(prodws.Cells[27, 3] as Excel.Range).Text;
                            }

                            wo.ProdStart = (string)(prodws.Cells[31, 1] as Excel.Range).Text;
                            wo.ProdEnd = (string)(prodws.Cells[31, 2] as Excel.Range).Text;
                            wo.ProdMin = (string)(prodws.Cells[31, 3] as Excel.Range).Text;
                            wo.FirstHour = (string)(prodws.Cells[35, 3] as Excel.Range).Text;

                            if ((prodws.Cells[41, 6] as Excel.Range).Text != null)
                            {
                                wo.Bot5 = (string)(prodws.Cells[41, 6] as Excel.Range).Text;
                            }
                            if ((prodws.Cells[41, 8] as Excel.Range).Text != null)
                            {
                                wo.BotTec = (string)(prodws.Cells[41, 8] as Excel.Range).Text;
                            }
                            if ((prodws.Cells[41, 9] as Excel.Range).Text != null)
                            {
                                wo.BotTemp = (string)(prodws.Cells[41, 9] as Excel.Range).Text;
                            }
                            if ((prodws.Cells[45, 9] as Excel.Range).Text != null)
                            {
                                wo.Lunch = (string)(prodws.Cells[45, 9] as Excel.Range).Text;
                            }

                            if ((prodws.Cells[46, 9] as Excel.Range).Text != null)
                            {
                                wo.Meeting = (string)(prodws.Cells[46, 9] as Excel.Range).Text;
                            }

                            if ((prodws.Cells[47, 9] as Excel.Range).Text != null)
                            {
                                wo.Brk = (string)(prodws.Cells[47, 9] as Excel.Range).Text;
                            }

                            if ((prodws.Cells[48, 9] as Excel.Range).Text != null)
                            {
                                wo.JobChange = (string)(prodws.Cells[48, 9] as Excel.Range).Text;
                            }

                            if ((prodws.Cells[49, 9] as Excel.Range).Text != null)
                            {
                                wo.Setup = (string)(prodws.Cells[49, 9] as Excel.Range).Text;
                            }

                            if ((prodws.Cells[50, 9] as Excel.Range).Text != null)
                            {
                                wo.Housekeeping = (string)(prodws.Cells[50, 9] as Excel.Range).Text;
                            }
//~~Downtime Read~~//
                            Excel._Worksheet dtws = wb.Sheets[2];
                            Excel.Range dtrng = prodws.UsedRange;

                            int rowCount = 42;
                            int colCount = 33;
                            int counter = 0;

                            string[,] dtValues = new string[42, 7];

                            for (int i = 7; i <= rowCount; i++)
                            {
                                for (int j = 1; j <= colCount; j++)
                                {
                                    if (j == 1 && dtrng.Cells[i, j].Text != null)
                                    {
                                        counter++;
                                        dtValues[counter - 1, 0] = wo.ProdDate; //Production Date
                                        dtValues[counter - 1, 1] = wo.Workorder; //Work Order
                                        dtValues[counter - 1, 2] = wo.Shift; //Shift
                                        dtValues[counter - 1, 3] = (string)dtrng.Cells[i, j].Text; //Column 1, Row 1 - The start time of the downtime event
                                    }
                                    else if (j == 33 && dtrng.Cells[i, j].Value != null)
                                    {
                                        dtValues[counter - 1, 6] = (string)dtrng.Cells[i, j].Text; //Downtime comment
                                    }
                                    else if (dtrng.Cells[i, j] != null && dtrng.Cells[i, j].Text != null)
                                    {
                                        dtValues[counter - 1, 4] = dtCategories[j]; //Downtime category
                                        dtValues[counter - 1, 5] = (string)dtrng.Cells[i, j].Text; //Downtime value
                                    }
                                }
                            }
}