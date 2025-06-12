using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Windows.Forms;
using Civil.Software;
using System.Windows;
using System.Collections.Generic;
using Autodesk.AutoCAD.GraphicsInterface;
using System.Windows.Shapes;
using AcadLine = Autodesk.AutoCAD.DatabaseServices.Line;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.Colors;
namespace Civil
{
    public class Initialization : IExtensionApplication
    {
        [CommandMethod("MyFirstLine")]
        public void cmdMyFirst()
        {
            Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
            // Prompt user to specify the insertion point
            PromptPointResult prPtRes1 = ed.GetPoint("\nSpecify Insertion point: ");
            if (prPtRes1.Status != PromptStatus.OK) return;
            Point3d pnt1 = prPtRes1.Value;
            // Show input form and retrieve user-defined values
            using (CC_Wall_Input form = new CC_Wall_Input())
            {
                if (form.ShowDialog() != DialogResult.OK) return;
                // Extract values from form inputs
                double St = form.St_o*1000;
                double H = form.H_o*1000;
                double Ft = form.Ft_o * 1000;
                double FB = form.FB_o * 1000;
                double RB = form.RB_o * 1000;
                double FOF = form.FOF_o * 1000;    
                double ROF = form.ROF_o * 1000;
                double FBH = form.FBH_o * 1000;
                double RBH = form.RBH_o * 1000;
                double rebar_m = (double)form.rebar_m_o;
                double spacing_m = (double)form.spacing_m_o;
                double rebar_d = (double)form.rebar_d_o;
                double spacing_d = (double)form.spacing_d_o;

                Database db = HostApplicationServices.WorkingDatabase;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord msBlkRec = trans.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    // Create a dimension style if needed
                    ObjectId dimStyleId = GetOrCreateDimStyle(db, trans);//Create a dimension style,  if needed
                    // Define key wall points
                    Point3d[] Pnts = new Point3d[10];
                    Pnts[0] = new Point3d(pnt1.X, pnt1.Y, 0);
                    Pnts[1] = new Point3d(pnt1.X - St, pnt1.Y, 0);
                    Pnts[2] = new Point3d(Pnts[1].X, pnt1.Y - (H - RBH), 0);
                    Pnts[3] = new Point3d(Pnts[2].X - RB, Pnts[2].Y - RBH, 0);
                    Pnts[4] = new Point3d(Pnts[3].X - ROF, Pnts[3].Y, 0);
                    Pnts[5] = new Point3d(Pnts[4].X, Pnts[4].Y - Ft, 0);

                    Pnts[9] = new Point3d(Pnts[0].X, pnt1.Y - (H - FBH), 0);
                    Pnts[8] = new Point3d(Pnts[9].X + FB, Pnts[9].Y - FBH, 0);
                    Pnts[7] = new Point3d(Pnts[8].X + FOF, Pnts[8].Y, 0);
                    Pnts[6] = new Point3d(Pnts[7].X, Pnts[7].Y - Ft, 0);

                    // Create wall outline by drawing lines between defined points
                    for (int i = 0; i < 10; i++)
                    {
                        AddLine(msBlkRec, trans, Pnts[i], Pnts[(i + 1) % 10],"Wall");
                    }



                    // Add dimensions for horizontal and vertical measurements
                    int dim_dist = 200;
                    if (RB != 0)
                    AddDimension(msBlkRec, trans, Pnts[0], Pnts[1], new Point3d((Pnts[0].X + Pnts[0].X) / 2, Pnts[0].Y + dim_dist, 0), dimStyleId, 0);

                    {
                        AddDimension(msBlkRec, trans, Pnts[3], Pnts[2], new Point3d((Pnts[3].X + Pnts[2].X) / 2, Pnts[3].Y + dim_dist, 0), dimStyleId, 0);
                    }
                    if (FB != 0)
                    {
                        AddDimension(msBlkRec, trans, Pnts[9], Pnts[8], new Point3d((Pnts[9].X + Pnts[8].X) / 2, Pnts[8].Y + dim_dist, 0), dimStyleId, 0);
                    }
                    AddDimension(msBlkRec, trans, Pnts[3], Pnts[4], new Point3d((Pnts[3].X + Pnts[4].X) / 2, Pnts[3].Y + dim_dist, 0), dimStyleId, 0);
                    AddDimension(msBlkRec, trans, Pnts[7], Pnts[8], new Point3d((Pnts[7].X + Pnts[8].X) / 2, Pnts[8].Y + dim_dist, 0), dimStyleId, 0);

                    AddDimension(msBlkRec, trans, Pnts[5], Pnts[6], new Point3d((Pnts[5].X + Pnts[6].X) / 2, Pnts[5].Y - dim_dist, 0), dimStyleId, 0);



                    //Vertical

                    AddDimension(msBlkRec, trans, Pnts[4], Pnts[1], new Point3d(Pnts[4].X - dim_dist, 0.5 * (Pnts[4].Y + Pnts[1].Y), 0), dimStyleId, Math.PI * 0.5);
                    AddDimension(msBlkRec, trans, Pnts[4], Pnts[5], new Point3d(Pnts[4].X - dim_dist, 0.5 * (Pnts[5].Y + Pnts[4].Y), 0), dimStyleId, Math.PI * 0.5);

                    if (RBH != 0 || RB != 0)
                    {
                        AddDimension(msBlkRec, trans, Pnts[1], Pnts[2], new Point3d(Pnts[2].X - dim_dist, 0.5 * (Pnts[1].Y + Pnts[2].Y), 0), dimStyleId, Math.PI * 0.5);
                        AddDimension(msBlkRec, trans, Pnts[2], Pnts[3], new Point3d(Pnts[2].X - dim_dist, 0.5 * (Pnts[2].Y + Pnts[3].Y), 0), dimStyleId, Math.PI * 0.5);
                    }
                    bool batter = false;
                    if (FBH != 0 || FB != 0)
                    {
                        AddDimension(msBlkRec, trans, Pnts[0], Pnts[9], new Point3d(Pnts[9].X + dim_dist, 0.5 * (Pnts[0].Y + Pnts[9].Y), 0), dimStyleId, Math.PI * 0.5);
                        AddDimension(msBlkRec, trans, Pnts[9], Pnts[8], new Point3d(Pnts[8].X + dim_dist, 0.5 * (Pnts[9].Y + Pnts[8].Y), 0), dimStyleId, Math.PI * 0.5);
                        batter = true;
                    }
                    // Add reinforcement elements
                    double topld = Math.Min(50 * rebar_d, St-150);
                    double botld = Math.Min(50 * rebar_d, Ft-150);

                    ObjectId rebarLayerId = GetOrCreateLayer(msBlkRec.Database, trans, "Reinforcement", 2);
                    //AddRebarCircle(msBlkRec, trans, Pnts[0], 0.2 / 2, rebarLayerId);
                    AddReinforcement(msBlkRec, trans, Pnts[0], Pnts[9], Pnts[8], 100,rebar_m, spacing_m, rebar_d, spacing_d, topld, botld, batter);
                     
                    trans.Commit();


                    
                }
            }
        }

        [CommandMethod("DrawWeirSection")]
        public void cmdDrawWeirSection()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            // Prompt user to browse for the Excel file
            string excelFilePath = "";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select Excel Workbook for Weir Section Data";
                openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|All Files (*.*)|*.*";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.RestoreDirectory = true;

                DialogResult result = openFileDialog.ShowDialog();
                if (result != DialogResult.OK)
                {
                    ed.WriteMessage("\nFile selection canceled by user.");
                    return;
                }
                excelFilePath = openFileDialog.FileName;
            }

            // Validate file existence
            if (!System.IO.File.Exists(excelFilePath))
            {
                ed.WriteMessage("\nExcel file not found at: " + excelFilePath);
                return;
            }

            // Initialize Excel application
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
            Microsoft.Office.Interop.Excel.Sheets sheets = null;
            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(excelFilePath);
                sheets = workbook.Sheets;

                // Get list of sheet names
                List<string> sheetNames = new List<string>();
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
                {
                    sheetNames.Add(sheet.Name);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet); // Release each sheet object
                }

                // Prompt user to select a sheet
                ed.WriteMessage("\nAvailable sheets in the Excel file:\n");
                for (int i = 0; i < sheetNames.Count; i++)
                {
                    ed.WriteMessage($"[{i + 1}] {sheetNames[i]}\n");
                }

                PromptIntegerOptions pio = new PromptIntegerOptions("\nSelect a sheet number (1-" + sheetNames.Count + "): ");
                pio.AllowZero = false;
                pio.AllowNegative = false;
                pio.LowerLimit = 1;
                pio.UpperLimit = sheetNames.Count;
                PromptIntegerResult pir = ed.GetInteger(pio);
                if (pir.Status != PromptStatus.OK) return;

                // Select the worksheet
                worksheet = workbook.Sheets[pir.Value];
                range = worksheet.UsedRange;

                // Read and validate dimensions from the Excel file
                double lengthWeir = 0.0, crestLevel = 0.0, depthHighCoeff = 0.0, widthStem = 0.0, usBatterWidth = 0.0,
                       dsBatterWidth = 0.0, foundationOffset = 0.0, foundationTopLevel = 0.0, foundationBottomLevel = 0.0,
                       Wcthick_body = 0.0, Wcthick_apron = 0.0, apron_length = 0.0, apron_thick = 0.0, apron_level = 0.0,
                       US_cutOff_level = 0.0, US_cutOff_slope = 0.0, US_cutOff_thickness = 0.0, cisternThickness = 0.0,
                       cisternLength = 0.0, endSillLevel = 0.0, endSillSlope = 0.0, endSillTopWidth = 0.0,
                       DS_cutOff_level = 0.0, DS_cutOff_slope = 0.0, DS_cutOff_thickness = 0.0,cisternLevel = 0.0;

                // Helper function to read and validate a cell
                bool ReadAndValidateCell(int row, int col, out double value, string paramName, bool mustBePositive = true)
                {
                    value = 0.0;
                    var cell = range.Cells[row, col];
                    bool isValid = cell.Value != null && double.TryParse(cell.Value.ToString(), out value) && (!mustBePositive || value > 0);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(cell); // Release cell COM object
                    if (!isValid)
                    {
                        ed.WriteMessage($"\nInvalid or missing {paramName} in Excel (cell B{row}). Must be a {(mustBePositive ? "positive number" : "number")}.");
                    }
                    return isValid;
                }

                // Read and validate each variable
                if (!ReadAndValidateCell(2, 2, out lengthWeir, "lengthWeir")) return; lengthWeir *= 1000;
                if (!ReadAndValidateCell(3, 2, out crestLevel, "crestLevel", false)) return; crestLevel *= 1000;
                if (!ReadAndValidateCell(4, 2, out depthHighCoeff, "depthHighCoeff", false)) return; depthHighCoeff *= 1000;
                if (!ReadAndValidateCell(5, 2, out widthStem, "widthStem")) return; widthStem *= 1000;
                if (!ReadAndValidateCell(6, 2, out usBatterWidth, "usBatterWidth", false)) return; usBatterWidth *= 1000;
                if (!ReadAndValidateCell(7, 2, out dsBatterWidth, "dsBatterWidth", false)) return; dsBatterWidth *= 1000;
                if (!ReadAndValidateCell(8, 2, out foundationOffset, "foundationOffset", false)) return; foundationOffset *= 1000;
                if (!ReadAndValidateCell(9, 2, out foundationTopLevel, "foundationTopLevel", false)) return; foundationTopLevel *= 1000;
                if (!ReadAndValidateCell(10, 2, out foundationBottomLevel, "foundationBottomLevel", false)) return; foundationBottomLevel *= 1000;
                if (!ReadAndValidateCell(11, 2, out Wcthick_body, "wearing coat thickness on body", false)) return; Wcthick_body *= 1000;
                if (!ReadAndValidateCell(12, 2, out apron_thick, "U/S apron thickness")) return; apron_thick *= 1000;
                if (!ReadAndValidateCell(13, 2, out apron_level, "U/S apron level", false)) return; apron_level *= 1000;
                if (!ReadAndValidateCell(14, 2, out Wcthick_apron, "wearing coat thickness on apron", false)) return; Wcthick_apron *= 1000;
                if (!ReadAndValidateCell(15, 2, out apron_length, "apron length")) return; apron_length *= 1000;
                if (!ReadAndValidateCell(16, 2, out US_cutOff_level, "U/S cut-off level", false)) return; US_cutOff_level *= 1000;
                if (!ReadAndValidateCell(17, 2, out US_cutOff_slope, "U/S cut-off slope")) return;
                if (!ReadAndValidateCell(18, 2, out US_cutOff_thickness, "U/S cut-off thickness")) return; US_cutOff_thickness *= 1000;
                if (!ReadAndValidateCell(19, 2, out cisternLevel, "cistern Level")) return; cisternLevel *= 1000;
                if (!ReadAndValidateCell(20, 2, out cisternThickness, "cistern thickness")) return; cisternThickness *= 1000;
                if (!ReadAndValidateCell(21, 2, out cisternLength, "cistern length")) return; cisternLength *= 1000;
                if (!ReadAndValidateCell(22, 2, out endSillLevel, "end sill level", false)) return; endSillLevel *= 1000;
                if (!ReadAndValidateCell(23, 2, out endSillSlope, "end sill slope")) return;
                if (!ReadAndValidateCell(24, 2, out endSillTopWidth, "end sill top width")) return; endSillTopWidth *= 1000;
                if (!ReadAndValidateCell(25, 2, out DS_cutOff_level, "D/S cut-off level", false)) return; DS_cutOff_level *= 1000;
                if (!ReadAndValidateCell(26, 2, out DS_cutOff_slope, "D/S cut-off slope")) return;
                if (!ReadAndValidateCell(27, 2, out DS_cutOff_thickness, "D/S cut-off thickness")) return; DS_cutOff_thickness *= 1000;

                // Validate level hierarchy
                if (!(crestLevel > foundationTopLevel && foundationTopLevel > foundationBottomLevel))
                {
                    ed.WriteMessage("\nInvalid level hierarchy: crestLevel ({0}) must be greater than foundationTopLevel ({1}), which must be greater than foundationBottomLevel ({2}).", crestLevel, foundationTopLevel, foundationBottomLevel);
                    return;
                }

                if (!(foundationTopLevel <= apron_level && apron_level >= foundationBottomLevel))
                {
                    ed.WriteMessage("\nInvalid apron level: apronLevel ({0}) must be between foundationTopLevel ({1}) and foundationBottomLevel ({2}).", apron_level, foundationTopLevel, foundationBottomLevel);
                    return;
                }

                if (!(foundationTopLevel <= endSillLevel && endSillLevel >= foundationBottomLevel))
                {
                    ed.WriteMessage("\nInvalid end sill level: endSillLevel ({0}) must be between foundationTopLevel ({1}) and foundationBottomLevel ({2}).", endSillLevel, foundationTopLevel, foundationBottomLevel);
                    return;
                }

                //if (!(foundationBottomLevel >= US_cutOff_level && US_cutOff_level >= foundationBottomLevel - (foundationTopLevel - foundationBottomLevel)))
                //{
                //    ed.WriteMessage("\nInvalid U/S cut-off level: US_cutOff_level ({0}) must be within a reasonable range below foundationBottomLevel ({1}).", US_cutOff_level, foundationBottomLevel);
                //    return;
                //}

                //if (!(foundationBottomLevel >= DS_cutOff_level && DS_cutOff_level >= foundationBottomLevel - (foundationTopLevel - foundationBottomLevel)))
                //{
                //    ed.WriteMessage("\nInvalid D/S cut-off level: DS_cutOff_level ({0}) must be within a reasonable range below foundationBottomLevel ({1}).", DS_cutOff_level, foundationBottomLevel);
                //    return;
                //}

                // Prompt user to specify insertion point
                PromptPointResult prPtRes1 = ed.GetPoint("\nSpecify Insertion point: ");
                if (prPtRes1.Status != PromptStatus.OK) return;
                Point3d pnt1 = prPtRes1.Value;

                // Calculate additional points for the weir section
                double crestBaseLevel = crestLevel - depthHighCoeff;
                double baseWidth = widthStem + usBatterWidth + dsBatterWidth;
                double foundationWidth = baseWidth + 2 * foundationOffset;
                double totalWidth = foundationWidth + apron_length + cisternLength;
                double xStart = pnt1.X;

                Database db = HostApplicationServices.WorkingDatabase;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord msBlkRec = trans.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    ObjectId dimStyleId = GetOrCreateDimStyle(db, trans);

                    // Define points for the weir section
                    List<Point3d> Pnts = new List<Point3d>(28); // Reserve space for indices 0–21
                    for (int i = 0; i < 29; i++)
                    {
                        Pnts.Add(Point3d.Origin); // Initialize with default points to avoid null access
                    }

                    // U/S cut-off points
                    double usCutOffDepth = apron_level - Wcthick_apron - apron_thick - US_cutOff_level;
                    double usCutOffWidth = usCutOffDepth /US_cutOff_slope + US_cutOff_thickness;
                    Pnts[1] = new Point3d(xStart, US_cutOff_level, 0); // P1: U/S cut-off bottom left
                    Pnts[2] = new Point3d(xStart + US_cutOff_thickness, US_cutOff_level, 0); // P2: U/S cut-off bottom right
                    Pnts[3] = new Point3d(xStart + usCutOffWidth, US_cutOff_level + usCutOffDepth, 0); // P3: U/S cut-off top right

                    // Apron points
                    Pnts[4] = new Point3d(xStart, apron_level, 0); // P4: Apron top left
                    Pnts[5] = new Point3d(xStart + apron_length+ (apron_level - foundationTopLevel) * usBatterWidth / (crestBaseLevel - foundationTopLevel), apron_level, 0); // P5: Apron top right

                    Pnts[9] = new Point3d(xStart + apron_length, foundationTopLevel, 0); // Body wall base left
                    Pnts[10] = new Point3d(Pnts[9].X - foundationOffset, foundationTopLevel, 0); // Foundation left
                    Pnts[11] = new Point3d(Pnts[10].X, foundationBottomLevel, 0);
                    Pnts[12] = new Point3d(Pnts[10].X + foundationWidth, foundationBottomLevel, 0); // Body wall base right
                    Pnts[21] = new Point3d(Pnts[12].X, foundationTopLevel, 0);
                    Pnts[20] = new Point3d(Pnts[21].X - foundationOffset, foundationTopLevel, 0);
                    Pnts[18] = new Point3d(Pnts[20].X - dsBatterWidth, crestLevel, 0);
                    Pnts[19] = new Point3d(Pnts[18].X - widthStem, crestBaseLevel, 0);

                   
                    // Body wall
                    AcadLine fTopLine = AddLine(msBlkRec, trans, Pnts[10], Pnts[21], "Wall"); // Base
                    AcadLine line2 = AddLine(msBlkRec, trans, Pnts[9], Pnts[19], "Wall"); // U/S batter
                    AcadLine crestLine = AddLine(msBlkRec, trans, Pnts[19], Pnts[18], "Wall"); // Crest
                    AcadLine dsline = AddLine(msBlkRec, trans, Pnts[18], Pnts[20], "Wall"); // D/S batter
                    AddLine(msBlkRec, trans, Pnts[10], Pnts[11], "Wall"); // Left thickness
                    AcadLine foundationRight = AddLine(msBlkRec, trans, Pnts[12], Pnts[21], "Wall"); // Right thickness
                    AddLine(msBlkRec, trans, Pnts[11], Pnts[12], "Wall"); // Foundation bottom
                                     // Wearing coat Offset 
                    AcadLine usLine1 = CreateOffsetLine(msBlkRec, trans, line2, -Wcthick_body, fTopLine, "Wall");
                    AcadLine crestLine1 = CreateOffsetLine(msBlkRec, trans, crestLine, -Wcthick_body, usLine1, "Wall");
                    AcadLine dsline1 = CreateOffsetLine(msBlkRec, trans, dsline, -Wcthick_body, crestLine1, "Wall");

                    usLine1 = TrimLine(usLine1, crestLine1, Pnts[5]);
                    crestLine1 = TrimLine(crestLine1, dsline1, Pnts[19]);
                    dsline1 = TrimLine(dsline1, fTopLine, Pnts[18]);

                    // Apron
                    AcadLine line1 = AddLine(msBlkRec, trans, Pnts[4], Pnts[5], "Wall"); // Apron top
                    AddLine(msBlkRec, trans, Pnts[1], Pnts[2], "Wall");//cut-off bottom
                    AddLine(msBlkRec, trans, Pnts[2], Pnts[3], "Wall");//cut-off slope
                    AddLine(msBlkRec, trans, Pnts[1], Pnts[4], "Wall");//left thick of apron
                    // Create offset lines for wearing coat and apron thickness
                    AcadLine line3 = CreateOffsetLine(msBlkRec, trans, line1, -Wcthick_apron, line2, "Wall");
                    if (line3 == null)
                    {
                        ed.WriteMessage("\nError: Failed to create wearing coat offset line.");
                        return;
                    }
                    AcadLine line4 = CreateOffsetLine(msBlkRec, trans, line3, apron_thick, line2, "Wall");
                    if (line4 == null)
                    {
                        ed.WriteMessage("\nError: Failed to create apron thickness offset line.");
                        return;
                    }
                    // Cistern and end sill points
                    Pnts[22] = new Point3d(Pnts[20].X + cisternLength, cisternLevel-Wcthick_body, 0);
                    Pnts[23] = new Point3d(Pnts[22].X + (endSillLevel - foundationTopLevel - Wcthick_body) * endSillSlope, endSillLevel - Wcthick_body, 0);

                    AcadLine cisBotLine = AddLine(msBlkRec, trans, new Point3d(Pnts[9].X, Pnts[22].Y, 0), Pnts[22], "Wall");
                   // cisBotLine.StartPoint= new Point3d(Pnts[9].X, Pnts[22].Y, 0);
                    if (GetIntersectionPoint(cisBotLine, dsline) != null)
                        cisBotLine = TrimLine(cisBotLine, dsline, Pnts[22]);
                    else if (GetIntersectionPoint(cisBotLine, foundationRight) != null)
                        cisBotLine = TrimLine(cisBotLine, foundationRight, Pnts[22]);
                    else
                        cisBotLine.StartPoint = new Point3d(Pnts[20].X, Pnts[22].Y, 0);
                    cisBotLine.Color = Color.FromRgb(0, 255, 0);

                    //cisBotLine = ExtendLine(cisBotLine, dsline1);//, Pnts[28]);
                    AcadLine cisTopLine = CreateOffsetLine(msBlkRec, trans, cisBotLine, Wcthick_body, null, "Wall");


                    AcadLine sillSlopeInLine = AddLine(msBlkRec, trans, Pnts[22], Pnts[23], "Wall");
                    AcadLine sillSlopeLine = CreateOffsetLine(msBlkRec, trans, sillSlopeInLine, Wcthick_body, null, "Wall");

                    cisTopLine = ExtendLine(cisTopLine, sillSlopeInLine);//, Pnts[20])
                    sillSlopeLine = ExtendLine(sillSlopeLine, cisBotLine);//, Pnts[23]);

                    // Pnts[16] = (Point3d)GetIntersectionPoint(cisTopLine, sillSlopeLine);
                    cisTopLine = TrimLine(cisTopLine, sillSlopeLine, Pnts[20]);
                    sillSlopeLine = TrimLine(sillSlopeLine,cisTopLine, Pnts[23]);

                    cisTopLine = ExtendLine(cisTopLine, dsline);

                    AcadLine sillTopLine = AddLine(msBlkRec, trans, new Point3d(Pnts[20].X, endSillLevel, 0), new Point3d(Pnts[23].X, endSillLevel, 0), "Wall");
                    
                    sillSlopeLine = ExtendLine(sillSlopeLine, sillTopLine);

                    sillTopLine = TrimLine(sillTopLine, sillSlopeLine, Pnts[23]);

                    double cisternBaseLevel = cisternLevel - cisternThickness - Wcthick_body;
                    Pnts[15] = (Point3d)GetIntersectionPoint(sillSlopeLine, sillTopLine);
                    Pnts[14] = new Point3d(Pnts[15].X + endSillTopWidth, Pnts[15].Y, 0);
                    Pnts[24] = new Point3d(Pnts[14].X , Pnts[14].Y-Wcthick_body, 0);
                    Pnts[13] = new Point3d(Pnts[14].X, cisternBaseLevel, 0);
                    Pnts[25] = new Point3d(Pnts[14].X, DS_cutOff_level, 0);
                    Pnts[26] = new Point3d(Pnts[25].X-DS_cutOff_thickness, DS_cutOff_level, 0);
                    Pnts[27] = new Point3d(Pnts[26].X - (cisternBaseLevel-DS_cutOff_level)/DS_cutOff_slope, cisternBaseLevel, 0);

                    AddLine(msBlkRec, trans, Pnts[25], Pnts[26], "Wall");
                    AddLine(msBlkRec, trans, Pnts[26], Pnts[27], "Wall");
                    AcadLine cisternBottomLine =AddLine(msBlkRec, trans, Pnts[13], new Point3d(Pnts[9].X, Pnts[13].Y, 0), "Wall");
                    AcadLine rightLine = AddLine(msBlkRec, trans, Pnts[14], Pnts[25], "Wall");
                    sillTopLine = ExtendLine(sillTopLine, rightLine);
                    if(GetIntersectionPoint(cisternBottomLine, dsline)!=null)
                        cisternBottomLine = TrimLine(cisternBottomLine, dsline, Pnts[13]);
                    else if(GetIntersectionPoint(cisternBottomLine, foundationRight) != null)
                        cisternBottomLine = TrimLine(cisternBottomLine, foundationRight, Pnts[13]);
                    else
                        cisternBottomLine.EndPoint = new Point3d(Pnts[20].X, Pnts[13].Y, 0);

                    CreateOffsetLine(msBlkRec, trans, sillTopLine, -Wcthick_body, sillSlopeInLine, "Wall");


                    // Commented code from original implementation
                    /*
                    // Foundation points
                    Pnts.Add(new Point3d(xStart, foundationBottomLevel, 0)); // P0: Foundation bottom left
                    Pnts.Add(new Point3d(xStart + totalWidth, foundationBottomLevel, 0)); // P1: Foundation bottom right
                    Pnts.Add(new Point3d(xStart, foundationTopLevel, 0)); // P2: Foundation top left
                    Pnts.Add(new Point3d(xStart + totalWidth, foundationTopLevel, 0)); // P3: Foundation top right

                    // Cistern and end sill points
                    Pnts.Add(new Point3d(xStart + foundationOffset + apron_length + baseWidth, endSillLevel, 0)); // P15: Cistern top left
                    Pnts.Add(new Point3d(xStart + foundationOffset + apron_length + baseWidth + cisternLength - endSillTopWidth, endSillLevel, 0)); // P16: End sill base left
                    Pnts.Add(new Point3d(xStart + foundationOffset + apron_length + baseWidth + cisternLength, endSillLevel, 0)); // P17: End sill top
                    Pnts.Add(new Point3d(xStart + foundationOffset + apron_length + baseWidth + cisternLength - endSillTopWidth + (endSillLevel - foundationTopLevel) * endSillSlope, foundationTopLevel, 0)); // P18: End sill base right
                    Pnts.Add(new Point3d(xStart + foundationOffset + apron_length + baseWidth, endSillLevel - cisternThickness, 0)); // P19: Cistern bottom left
                    Pnts.Add(new Point3d(xStart + foundationOffset + apron_length + baseWidth + cisternLength - endSillTopWidth, endSillLevel - cisternThickness, 0)); // P20: Cistern bottom right at end sill base

                    // D/S cut-off points
                    double dsCutOffDepth = foundationBottomLevel - DS_cutOff_level;
                    double dsCutOffWidth = dsCutOffDepth * DS_cutOff_slope;
                    Pnts.Add(new Point3d(xStart + totalWidth - foundationOffset, foundationTopLevel, 0)); // P21: D/S cut-off top
                    Pnts.Add(new Point3d(xStart + totalWidth - foundationOffset + dsCutOffWidth - DS_cutOff_thickness, DS_cutOff_level, 0)); // P22: D/S cut-off bottom left
                    Pnts.Add(new Point3d(xStart + totalWidth - foundationOffset + dsCutOffWidth, DS_cutOff_level, 0)); // P23: D/S cut-off bottom right

                    // Draw the weir section components
                    // Foundation
                    AddLine(msBlkRec, trans, Pnts[0], Pnts[1], "Wall"); // Bottom
                    AddLine(msBlkRec, trans, Pnts[1], Pnts[3], "Wall"); // Right
                    AddLine(msBlkRec, trans, Pnts[2], Pnts[3], "Wall"); // Top
                    AddLine(msBlkRec, trans, Pnts[0], Pnts[2], "Wall"); // Left

                    // U/S cut-off
                    AddLine(msBlkRec, trans, Pnts[4], Pnts[5], "Wall"); // Left side
                    AddLine(msBlkRec, trans, Pnts[4], Pnts[6], "Wall"); // Bottom
                    AddLine(msBlkRec, trans, Pnts[6], Pnts[5], "Wall"); // Right side

                    // Apron
                    AddLine(msBlkRec, trans, Pnts[7], Pnts[8], "Wall"); // Top
                    AddLine(msBlkRec, trans, Pnts[9], Pnts[10], "Wall"); // Bottom
                    AddLine(msBlkRec, trans, Pnts[7], Pnts[9], "Wall"); // Left
                    AddLine(msBlkRec, trans, Pnts[8], Pnts[10], "Wall"); // Right

                    // Cistern and end sill
                    AddLine(msBlkRec, trans, Pnts[15], Pnts[16], "Wall"); // Cistern top
                    AddLine(msBlkRec, trans, Pnts[19], Pnts[20], "Wall"); // Cistern bottom
                    AddLine(msBlkRec, trans, Pnts[15], Pnts[19], "Wall"); // Cistern left
                    AddLine(msBlkRec, trans, Pnts[16], Pnts[20], "Wall"); // Cistern right at end sill base
                    AddLine(msBlkRec, trans, Pnts[16], Pnts[17], "Wall"); // End sill left
                    AddLine(msBlkRec, trans, Pnts[17], Pnts[18], "Wall"); // End sill right

                    // D/S cut-off
                    AddLine(msBlkRec, trans, Pnts[21], Pnts[22], "Wall"); // Left side
                    AddLine(msBlkRec, trans, Pnts[22], Pnts[23], "Wall"); // Bottom
                    AddLine(msBlkRec, trans, Pnts[23], Pnts[21], "Wall"); // Right side

                    // Draw wearing coat offset lines
                    // Body wall wearing coat (inner offset)
                    List<Line> bodyWallLines = new List<Line>
                    {
                        new Line(Pnts[11], Pnts[13]),
                        new Line(Pnts[13], Pnts[14]),
                        new Line(Pnts[14], Pnts[12])
                    };
                    foreach (var line in bodyWallLines)
                    {
                        msBlkRec.AppendEntity(line);
                        trans.AddNewlyCreatedDBObject(line, true);
                        CreateOffsetLine(msBlkRec, trans, line, -Wcthick_body, null, "WearingCoat");
                    }

                    // Apron wearing coat (top offset)
                    Line apronTopLine = new Line(Pnts[7], Pnts[8]);
                    msBlkRec.AppendEntity(apronTopLine);
                    trans.AddNewlyCreatedDBObject(apronTopLine, true);
                    CreateOffsetLine(msBlkRec, trans, apronTopLine, Wcthick_apron, null, "WearingCoat");

                    // Cistern wearing coat (top offset)
                    Line cisternTopLine = new Line(Pnts[15], Pnts[16]);
                    msBlkRec.AppendEntity(cisternTopLine);
                    trans.AddNewlyCreatedDBObject(cisternTopLine, true);
                    CreateOffsetLine(msBlkRec, trans, cisternTopLine, Wcthick_body, null, "WearingCoat");

                    // Add dimensions
                    int dim_dist = 500;
                    // Horizontal dimensions
                    AddDimension(msBlkRec, trans, Pnts[0], Pnts[1], new Point3d((Pnts[0].X + Pnts[1].X) / 2, Pnts[0].Y - dim_dist, 0), dimStyleId, 0); // Foundation bottom
                    AddDimension(msBlkRec, trans, Pnts[4], Pnts[6], new Point3d((Pnts[4].X + Pnts[6].X) / 2, Pnts[4].Y - dim_dist / 2, 0), dimStyleId, 0); // U/S cut-off
                    AddDimension(msBlkRec, trans, Pnts[7], Pnts[8], new Point3d((Pnts[7].X + Pnts[8].X) / 2, Pnts[7].Y + dim_dist, 0), dimStyleId, 0); // Apron
                    AddDimension(msBlkRec, trans, Pnts[11], Pnts[12], new Point3d((Pnts[11].X + Pnts[12].X) / 2, Pnts[11].Y + dim_dist, 0), dimStyleId, 0); // Body wall base
                    AddDimension(msBlkRec, trans, Pnts[13], Pnts[14], new Point3d((Pnts[13].X + Pnts[14].X) / 2, Pnts[13].Y + dim_dist, 0), dimStyleId, 0); // Crest
                    AddDimension(msBlkRec, trans, Pnts[15], Pnts[16], new Point3d((Pnts[15].X + Pnts[16].X) / 2, Pnts[15].Y + dim_dist, 0), dimStyleId, 0); // Cistern
                    AddDimension(msBlkRec, trans, Pnts[16], Pnts[18], new Point3d((Pnts[16].X + Pnts[18].X) / 2, Pnts[16].Y + dim_dist, 0), dimStyleId, 0); // End sill base
                    AddDimension(msBlkRec, trans, Pnts[22], Pnts[23], new Point3d((Pnts[22].X + Pnts[23].X) / 2, Pnts[22].Y - dim_dist / 2, 0), dimStyleId, 0); // D/S cut-off

                    // Vertical dimensions
                    AddDimension(msBlkRec, trans, Pnts[0], Pnts[2], new Point3d(Pnts[0].X - dim_dist, (Pnts[0].Y + Pnts[2].Y) / 2, 0), dimStyleId, Math.PI * 0.5); // Foundation height
                    AddDimension(msBlkRec, trans, Pnts[9], Pnts[7], new Point3d(Pnts[9].X - dim_dist, (Pnts[9].Y + Pnts[7].Y) / 2, 0), dimStyleId, Math.PI * 0.5); // Apron thickness
                    AddDimension(msBlkRec, trans, Pnts[11], Pnts[13], new Point3d(Pnts[11].X - dim_dist, (Pnts[11].Y + Pnts[13].Y) / 2, 0), dimStyleId, Math.PI * 0.5); // U/S batter height
                    AddDimension(msBlkRec, trans, Pnts[13], Pnts[14], new Point3d(Pnts[13].X - dim_dist, (Pnts[13].Y + Pnts[14].Y) / 2, 0), dimStyleId, Math.PI * 0.5); // Crest height
                    AddDimension(msBlkRec, trans, Pnts[19], Pnts[15], new Point3d(Pnts[19].X + dim_dist, (Pnts[19].Y + Pnts[15].Y) / 2, 0), dimStyleId, Math.PI * 0.5); // Cistern thickness

                    // Add annotations
                    AddLeaderWithText(msBlkRec, trans, new Point3d((Pnts[13].X + Pnts[14].X) / 2, crestLevel, 0), $"CREST @ {crestLevel}");
                    AddLeaderWithText(msBlkRec, trans, new Point3d(Pnts[2].X, foundationTopLevel, 0), $"FOUNDATION TOP @ {foundationTopLevel}");
                    AddLeaderWithText(msBlkRec, trans, new Point3d(Pnts[0].X, foundationBottomLevel, 0), $"FOUNDATION BOTTOM @ {foundationBottomLevel}");
                    AddLeaderWithText(msBlkRec, trans, new Point3d(Pnts[7].X, apron_level, 0), $"U/S APRON @ {apron_level}");
                    AddLeaderWithText(msBlkRec, trans, new Point3d(Pnts[15].X, endSillLevel, 0), $"END SILL @ {endSillLevel}");
                    AddLeaderWithText(msBlkRec, trans, new Point3d(Pnts[4].X, US_cutOff_level, 0), $"U/S CUT-OFF @ {US_cutOff_level}");
                    AddLeaderWithText(msBlkRec, trans, new Point3d(Pnts[23].X, DS_cutOff_level, 0), $"D/S CUT-OFF @ {DS_cutOff_level}");
                    AddLeaderWithText(msBlkRec, trans, new Point3d((Pnts[13].X + Pnts[14].X) / 2, crestBaseLevel - 300, 0), $"{Wcthick_body} THICK WEARING COAT IN CC M20 GRADE");
                    AddLeaderWithText(msBlkRec, trans, new Point3d((Pnts[7].X + Pnts[8].X) / 2, apron_level + 300, 0), $"{Wcthick_apron} THICK WEARING COAT IN CC M20 GRADE");
                    AddLeaderWithText(msBlkRec, trans, new Point3d((Pnts[15].X + Pnts[16].X) / 2, endSillLevel + 300, 0), $"{Wcthick_body} THICK WEARING COAT IN CC M20 GRADE");
                    */

                    trans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError reading Excel file or drawing section: {ex.Message}");
            }
            finally
            {
                // Release COM objects
                if (range != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (sheets != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                // Force garbage collection to release COM references
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        [CommandMethod("DRAWREBAR")]
        public void DrawRebarCommand()
        {
            var ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            var db = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.WorkingDatabase;

            try
            {
                // Prompt user to select a line
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect a line for reinforcement: ");
                peo.SetRejectMessage("\nPlease select a line.");
                peo.AddAllowedClass(typeof(AcadLine), true);
                PromptEntityResult per = ed.GetEntity(peo);

                if (per.Status != PromptStatus.OK)
                    return;

                // Prompt for direction (left or right)
                PromptKeywordOptions pko = new PromptKeywordOptions("\nSpecify reinforcement direction [Left/Right]: ");
                pko.Keywords.Add("Left");
                pko.Keywords.Add("Right");
                pko.AllowNone = false;
                PromptResult pr = ed.GetKeywords(pko);

                if (pr.Status != PromptStatus.OK)
                    return;

                bool isRightSide = pr.StringResult == "Right";


                // Prompt for distribution bar diameter
                PromptDoubleOptions pdo = new PromptDoubleOptions("\nEnter distribution bar diameter (mm): ")
                {
                    AllowZero = false,
                    AllowNegative = false,
                    DefaultValue = 12.0
                };
                PromptDoubleResult pdr = ed.GetDouble(pdo);

                if (pdr.Status != PromptStatus.OK)
                    return;


                // Define reinforcement parameters
                double offset = 50.0;        // mm
                double barDiameter_m = 20.0; // Main bar diameter (mm)
                double spacing_m = 200.0;    // Main bar spacing (mm)
                double barDiameter_d = pdr.Value; ; // Distribution bar diameter (mm)
                double spacing_d = 150.0;    // Distribution bar spacing (mm)
                //double topld = 0;        // Top development length (mm)
               // double botld = 0;        // Bottom development length (mm)
                //bool batter = false;         // No batter for simplicity

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord msBlkRec = trans.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // Get the selected line
                    AcadLine selectedLine = (AcadLine)trans.GetObject(per.ObjectId, OpenMode.ForRead);
                    Point3d start = selectedLine.StartPoint;
                    Point3d end = selectedLine.EndPoint;

                    // Calculate mid-point for the third point (pt9)
                    Point3d midPoint = new Point3d(
                        (start.X + end.X) / 2.0,
                        (start.Y + end.Y) / 2.0,
                        0);

                    // Compute direction and perpendicular vectors
                    Vector3d direction = (end - start).GetNormal();
                    Vector3d perpDirection = new Vector3d(-direction.Y, direction.X, 0).GetNormal();
                    if (isRightSide)
                        perpDirection = -perpDirection; // Reverse for right side

                    // Call AddReinforcement with adjusted points
                    AddReinforcementForLine(
                        msBlkRec, trans,
                        start, end,
                        offset, barDiameter_m, spacing_m,
                        barDiameter_d, spacing_d,
                        perpDirection,isRightSide);

                    trans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        // Helper method to check if a point is outside the polygon
        private bool IsPointOutsidePolygon(Point3d point, Point3d[] polygon)
        {
            int n = polygon.Length;
            int crossings = 0;

            for (int i = 0; i < n; i++)
            {
                Point3d p1 = polygon[i];
                Point3d p2 = polygon[(i + 1) % n];

                if (((p1.Y <= point.Y && point.Y < p2.Y) || (p2.Y <= point.Y && point.Y < p1.Y)) &&
                    (point.X < (p2.X - p1.X) * (point.Y - p1.Y) / (p2.Y - p1.Y) + p1.X))
                {
                    crossings++;
                }
            }

            // Point is inside if number of crossings is odd
            return (crossings % 2 == 0);
        }

        // Creates or retrieves a predefined dimension style
        private ObjectId GetOrCreateDimStyle(Database db, Transaction trans)
        {
            DimStyleTable dimStyleTable = trans.GetObject(db.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
            string dimStyleName = "MyDimStyle";

            if (dimStyleTable.Has(dimStyleName))
            {
                return dimStyleTable[dimStyleName];
            }

            dimStyleTable.UpgradeOpen();
            DimStyleTableRecord dimStyle = new DimStyleTableRecord
            {
                Name = dimStyleName,
                Dimtxt = 200,
                Dimasz = 200,
                Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1),
                Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 2),
                Dimdec = 0,  // No decimal places
                Dimrnd = 1,  // Rounds to the nearest whole number
                Dimtoh = false
            };

            dimStyle.Dimltex1 = GetOrCreateLinetype(db, trans, "DASHED");
            dimStyle.Dimltex2 = GetOrCreateLinetype(db, trans, "DASHED");


            // dimStyle.Dimse1 = true;  // Suppress first extension line from the object
            // dimStyle.Dimse2 = true;  // Suppress second extension line from the object
            //imStyle.Dimfxl = 100;
            dimStyle.Dimfxlen = 100;
            dimStyle.Dimexe = 100;   // Set fixed length of extension lines at arrows
            dimStyle.Dimexo = 0;     // No offset from the object
            dimStyle.DimfxlenOn = true;  // Fixed length extension lines



            ObjectId dimStyleId = dimStyleTable.Add(dimStyle);
            trans.AddNewlyCreatedDBObject(dimStyle, true);
            return dimStyleId;
        }
        // Creates or retrieves a predefined linetype
        private ObjectId GetOrCreateLinetype(Database db, Transaction trans, string linetypeName)
        {
            LinetypeTable linetypeTable = trans.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

            if (linetypeTable.Has(linetypeName))
            {
                return linetypeTable[linetypeName];
            }

            db.LoadLineTypeFile(linetypeName, "acad.lin");

            return linetypeTable[linetypeName];
        }
        // Adds a dimension between two points
        private void AddDimension(BlockTableRecord msBlkRec, Transaction trans, Point3d pt1, Point3d pt2, Point3d dimLinePos, ObjectId dimStyleId, double rot)
        {
            RotatedDimension dim = new RotatedDimension(rot, pt1, pt2, dimLinePos, "", dimStyleId);
            ObjectId dimLayerId = GetOrCreateLayer(msBlkRec.Database, trans, "Dimensions", 3);
            dim.LayerId = dimLayerId;
            if (Math.Abs(rot - Math.PI * 0.5) < 0.01) // If it's near 90 degrees
            {
                dim.TextRotation = Math.PI * 0.5; // Ensures vertical text
                dim.Dimtad = 4;  // Centers text properly
            }

            msBlkRec.AppendEntity(dim);
            trans.AddNewlyCreatedDBObject(dim, true);
        }
        // Adds a line representing a wall
        private AcadLine AddLine(BlockTableRecord msBlkRec, Transaction trans, Point3d start, Point3d end,String layer)
        {
            AcadLine line = new AcadLine(start, end);
            ObjectId wallLayerId = GetOrCreateLayer(msBlkRec.Database, trans, layer, 1);
            line.LayerId = wallLayerId;
            msBlkRec.AppendEntity(line);
            trans.AddNewlyCreatedDBObject(line, true);
            return line;
        }
        // Creates or retrieves a layer
        private ObjectId GetOrCreateLayer(Database db, Transaction trans, string layerName, short colorIndex)
        {
            LayerTable layerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
            if (layerTable.Has(layerName))
            {
                return layerTable[layerName];
            }

            layerTable.UpgradeOpen();
            LayerTableRecord newLayer = new LayerTableRecord
            {
                Name = layerName,
                Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex)
            };

            ObjectId layerId = layerTable.Add(newLayer);
            trans.AddNewlyCreatedDBObject(newLayer, true);
            return layerId;
        }
        //private void AddLeaderWithText(BlockTableRecord msBlkRec, Transaction trans, Point3d attachPoint, string text)
        //{
        //    ObjectId leaderLayerId = GetOrCreateLayer(msBlkRec.Database, trans, "Annotations", 4);
        //    double dimasz = (double)acadApp.GetSystemVariable("DIMASZ");
        //    using (Leader leader = new Leader())
        //    {
        //        leader.SetDatabaseDefaults();
        //        leader.LayerId = leaderLayerId;
        //        leader.Dimasz = 200;
        //        leader.ColorIndex = 4; // Cyan for annotations
        //        leader.AppendVertex(attachPoint);
        //        //leader.AppendVertex(new Point3d(attachPoint.X + 300, attachPoint.Y + 300, 0)); // Leader elbow
        //        leader.AppendVertex(new Point3d(attachPoint.X + 700, attachPoint.Y + 400, 0)); // Final point
        //        msBlkRec.AppendEntity(leader);
        //        trans.AddNewlyCreatedDBObject(leader, true);
        //        // Attach multiline text to leader
        //        using (MText mtext = new MText())
        //        {
        //            mtext.SetDatabaseDefaults();
        //            mtext.Contents = text;
        //            mtext.Location = new Point3d(attachPoint.X + 700, attachPoint.Y + 400, 0);
        //            mtext.TextHeight = 150;
        //            mtext.LayerId = leaderLayerId;

        //            msBlkRec.AppendEntity(mtext);
        //            trans.AddNewlyCreatedDBObject(mtext, true);
        //            leader.Annotation = mtext.ObjectId;
        //            leader.Annotative = AnnotativeStates.True;
        //            leader.EvaluateLeader();
        //        }


        //    }
        //}


        private void AddLeaderWithText(BlockTableRecord msBlkRec, Transaction trans, Point3d attachPoint, string text,
             double offsetX = 700.0, double offsetY = 400.0, bool isFlipped = false)
        {
            ObjectId leaderLayerId = GetOrCreateLayer(msBlkRec.Database, trans, "Annotations.AutoCAD", 4);
            using (Leader leader = new Leader())
            {
                leader.SetDatabaseDefaults();
                leader.LayerId = leaderLayerId;
                leader.Dimasz = 200; // Arrow size
                leader.HasArrowHead = true; // Explicitly enable arrow
                leader.ColorIndex = 4; // Cyan for annotations

                Point3d textPoint = new Point3d(attachPoint.X + offsetX, attachPoint.Y + offsetY, 0);

                if (isFlipped)
                {
                    // Leader starts at text (top-right) and ends at attachPoint
                    leader.AppendVertex(textPoint);
                    leader.AppendVertex(attachPoint);
                }
                else
                {
                    // Leader starts at attachPoint and ends at text
                    leader.AppendVertex(attachPoint);
                    leader.AppendVertex(textPoint);
                }

                msBlkRec.AppendEntity(leader);
                trans.AddNewlyCreatedDBObject(leader, true);

                using (MText mtext = new MText())
                {
                    mtext.SetDatabaseDefaults();
                    mtext.Contents = text;
                    mtext.TextHeight = 150;
                    mtext.LayerId = leaderLayerId;

                    if (isFlipped)
                    {
                        mtext.Attachment = AttachmentPoint.TopRight;
                        mtext.Location = textPoint; // Top-right corner at textPoint
                    }
                    else
                    {
                        mtext.Attachment = AttachmentPoint.BottomLeft;
                        mtext.Location = textPoint; // Bottom-left corner at textPoint
                    }

                    msBlkRec.AppendEntity(mtext);
                    trans.AddNewlyCreatedDBObject(mtext, true);
                    leader.Annotation = mtext.ObjectId;
                    leader.Annotative = AnnotativeStates.True;
                    leader.EvaluateLeader();
                }
            }
        }



        // Adds reinforcement lines and rebar circles
        private void AddReinforcement(BlockTableRecord msBlkRec, Transaction trans, Point3d pt0, Point3d pt9, Point3d pt8, double offset, double barDiameter_m, double spacing_m, double barDiameter_d, double spacing_d,double topld, double botld, Boolean batter)
        {
            ObjectId rebarLayerId = GetOrCreateLayer(msBlkRec.Database, trans, "Reinforcement", 2);

            // Offset reinforcement path
            Point3d offsetPt0 = new Point3d(pt0.X - offset, pt0.Y - offset, 0);
            Point3d offsetPt9 = new Point3d(pt9.X - offset, pt9.Y - offset, 0);
            Point3d offsetPt8 = new Point3d(pt8.X - offset, pt8.Y - offset, 0);
           
            Point3d topbarpnt = new Point3d(offsetPt0.X - topld, offsetPt0.Y, 0);
            Point3d botbarpnt = new Point3d(offsetPt8.X, offsetPt8.Y-botld, 0);
            if (offsetPt0.DistanceTo(topbarpnt) < 50 * barDiameter_d)
            {
                Point3d ldpnt = new Point3d(topbarpnt.X, topbarpnt.Y - 30 * barDiameter_d, 0);
                AddLine(msBlkRec, trans, topbarpnt, ldpnt, "Reinforcement");
            }
            if (offsetPt8.DistanceTo(botbarpnt) < 50 * barDiameter_d)
            {
                Point3d ldpnt = new Point3d(botbarpnt.X - 30 * barDiameter_d, botbarpnt.Y, 0);
                AddLine(msBlkRec, trans, botbarpnt, ldpnt, "Reinforcement");
            }

                // Add offset reinforcement lines
            AddLine(msBlkRec, trans, offsetPt0, offsetPt9,"Reinforcement");
            AddLine(msBlkRec, trans, offsetPt9, offsetPt8, "Reinforcement");
            AddLine(msBlkRec, trans, offsetPt0, topbarpnt, "Reinforcement"); 
            AddLine(msBlkRec, trans, botbarpnt, offsetPt8, "Reinforcement");

            // Compute unit vector along reinforcement line
            Vector3d direction = (offsetPt9 - offsetPt0).GetNormal();
            //
            if (batter)
            {
                Point3d offsetPt_bar = offsetPt9 + (direction * 200);
                AddLine(msBlkRec, trans, offsetPt_bar, offsetPt9, "Reinforcement");
            }
            // Get perpendicular vector (normal to reinforcement line)
            Vector3d perpDirection = new Vector3d(-direction.Y, direction.X, 0).GetNormal();

            // Compute number of rebar circles
            double length = offsetPt0.DistanceTo(offsetPt9);
            int numBars = (int)(length / spacing_d);
            offsetPt0 = new Point3d(offsetPt0.X,offsetPt0.Y - 0.5* barDiameter_d,0);
            for (int i = 0; i <= numBars; i++)
            {
                // Move along the offset reinforcement line
                Point3d baseCenter = offsetPt0 + (direction * (i * spacing_d));

                // Shift the center outward by rebar radius (tangent positioning)
                Point3d barCenter = baseCenter - (perpDirection * (barDiameter_d / 2));

                // Create rebar circle
                AddRebarCircle(msBlkRec, trans, barCenter, barDiameter_d / 2, rebarLayerId);
            }
            //Add reinforcement description
            string reinforcementDescription = $"Ø{barDiameter_d} @ {spacing_d}mm";

            Point3d attachPoint = offsetPt0 + (direction * (3 * spacing_d));

            attachPoint = attachPoint - (perpDirection * (barDiameter_d / 2));

            AddLeaderWithText(msBlkRec, trans, attachPoint, reinforcementDescription);
            //main bar description

            reinforcementDescription = $"Ø{barDiameter_m} @ {spacing_m}mm";

            attachPoint = offsetPt0 + (direction * (4 * spacing_d));

            attachPoint = new Point3d(attachPoint.X, attachPoint.Y - spacing_d*0.5, 0);   

            //attachPoint = attachPoint + (perpDirection * (barDiameter_d / 2));

            AddLeaderWithText(msBlkRec, trans, attachPoint, reinforcementDescription);



            //Add reinforcement after front batter height

            direction = (offsetPt8 - offsetPt9).GetNormal();
            if (batter)
            {
                Point3d offsetPt_bar = offsetPt9 - (direction * 200);
                AddLine(msBlkRec, trans, offsetPt_bar, offsetPt9, "Reinforcement");
            }

            perpDirection = new Vector3d(-direction.Y, direction.X, 0).GetNormal();

            length = offsetPt9.DistanceTo(offsetPt8);

             numBars = (int)(length / spacing_d);

            for (int i = 0; i <= numBars; i++)
            {
                // Move along the offset reinforcement line
                Point3d baseCenter = offsetPt9 + (direction * (i * spacing_d));

                // Shift the center outward by rebar radius (tangent positioning)
                Point3d barCenter = baseCenter - (perpDirection * (barDiameter_d / 2));

                // Create rebar circle
                AddRebarCircle(msBlkRec, trans, barCenter, barDiameter_d / 2, rebarLayerId);
            }
        }
        private void AddReinforcementForLine(BlockTableRecord msBlkRec, Transaction trans, Point3d pt0, Point3d pt8,
    double offset, double barDiameter_m, double spacing_m, double barDiameter_d, double spacing_d,
    Vector3d perpDirection, bool isRightSide)
        {
            ObjectId rebarLayerId = GetOrCreateLayer(msBlkRec.Database, trans, "Reinforcement", 2);

            // Offset reinforcement path
            Point3d offsetPt0 = pt0 + (perpDirection * offset);
            Point3d offsetPt8 = pt8 + (perpDirection * offset);

            // Compute direction vector
            Vector3d direction = (offsetPt8 - offsetPt0).GetNormal();

            // Add reinforcement line
            AddLine(msBlkRec, trans, offsetPt0, offsetPt8, "Reinforcement");

            // Place rebars starting at least one diameter inside
            double d = barDiameter_d;
            double totalLength = offsetPt0.DistanceTo(offsetPt8);
            while (d < totalLength - barDiameter_d)
            {
                Point3d baseCenter = offsetPt0 + (direction * d);
                Point3d barCenter = baseCenter - (perpDirection * (barDiameter_d / 2));
                AddRebarCircle(msBlkRec, trans, barCenter, barDiameter_d / 2, rebarLayerId);
                d += spacing_d;
            }
            // Text height from AddLeaderWithText
            double textHeight = 150.0;


            //double textHeight = 150.0;

            // Determine leader offsets based on line slope and side
            bool isFlipped = direction.X > 0; // Flip for right-sloping lines
            double offsetMagnitude = Math.Sqrt(700.0 * 700.0 + 400.0 * 400.0); // ~806.23 units
            // Text opposite rebars: rebars at -perpDirection, text at +perpDirection
            Vector3d textOffsetDirection = isRightSide ? -perpDirection : perpDirection;
            if (isFlipped)
                textOffsetDirection = isRightSide ? perpDirection : -perpDirection; // Adjust for flipped leader
                                                                                    // Fix for CS0029: Cannot implicitly convert type 'Autodesk.AutoCAD.Geometry.Vector3d' to 'Autodesk.AutoCAD.Geometry.Point3d'

            // Original problematic line
            // Point3d textOffset = textOffsetDirection * offsetMagnitude;

            // Fixed line
            Point3d textOffset = new Point3d(
               textOffsetDirection.X * offsetMagnitude,
               textOffsetDirection.Y * offsetMagnitude,
               textOffsetDirection.Z * offsetMagnitude
            );

            // Find nearest distribution bar center at 0.25 * totalLength
            double distBarPos = 0.25 * totalLength;
            double nearestBarDist = Math.Round((distBarPos - barDiameter_d) / spacing_d) * spacing_d + barDiameter_d;
            Point3d distBarCenter = offsetPt0 + (direction * nearestBarDist) - (perpDirection * (barDiameter_d / 2));

            // Add reinforcement description for distribution bars at 0.25 * totalLength
            string reinforcementDescription = $"Ø{barDiameter_d} @ {spacing_d}mm";
            // Fix for CS0019: Operator '+' cannot be applied to operands of type 'Point3d' and 'Point3d'

            // The issue occurs because `Point3d` does not support direct addition of two `Point3d` objects.
            // Instead, you can use the `Add` method or work with `Vector3d` for vector arithmetic.

            Point3d distTextPoint = distBarCenter.Add(new Vector3d(textOffset.X, textOffset.Y, textOffset.Z));
            AddLeaderWithText(msBlkRec, trans, distBarCenter, reinforcementDescription, distTextPoint, isFlipped);

            // Add reinforcement description for main bars at 0.75 * totalLength
            reinforcementDescription = $"Ø{barDiameter_m} @ {spacing_m}mm";
            Point3d mainBarPoint = offsetPt0 + (direction * (0.75 * totalLength));
            // Fix for CS0019: Operator '+' cannot be applied to operands of type 'Point3d' and 'Point3d'

            // Original problematic line
            // Point3d mainTextPoint = mainBarPoint + textOffset + (textOffsetDirection * textHeight); // Vertical separation

            // Fixed line
            Point3d mainTextPoint = mainBarPoint.Add(new Vector3d(textOffset.X, textOffset.Y, textOffset.Z))
                                                .Add(textOffsetDirection.MultiplyBy(textHeight));
            AddLeaderWithText(msBlkRec, trans, mainBarPoint, reinforcementDescription, mainTextPoint, isFlipped);
        }

        private void AddLeaderWithText(BlockTableRecord msBlkRec, Transaction trans, Point3d attachPoint, string text,
            Point3d textPoint, bool isFlipped = false)
        {
            ObjectId leaderLayerId = GetOrCreateLayer(msBlkRec.Database, trans, "Annotations", 4);
            using (Leader leader = new Leader())
            {
                leader.SetDatabaseDefaults();
                leader.LayerId = leaderLayerId;
                leader.Dimasz = 200; // Arrow size
                leader.HasArrowHead = true; // Explicitly enable arrow
                leader.ColorIndex = 4; // Cyan for annotations

                if (isFlipped)
                {
                    // Leader starts at text (top-right) and ends at attachPoint
                    leader.AppendVertex(textPoint);
                    leader.AppendVertex(attachPoint);
                }
                else
                {
                    // Leader starts at attachPoint and ends at text
                    leader.AppendVertex(attachPoint);
                    leader.AppendVertex(textPoint);
                }

                msBlkRec.AppendEntity(leader);
                trans.AddNewlyCreatedDBObject(leader, true);

                using (MText mtext = new MText())
                {
                    mtext.SetDatabaseDefaults();
                    mtext.Contents = text;
                    mtext.TextHeight = 150;
                    mtext.LayerId = leaderLayerId;

                    if (isFlipped)
                    {
                        mtext.Attachment = AttachmentPoint.TopRight;
                        mtext.Location = textPoint; // Top-right corner at textPoint
                    }
                    else
                    {
                        mtext.Attachment = AttachmentPoint.BottomLeft;
                        mtext.Location = textPoint; // Bottom-left corner at textPoint
                    }

                    msBlkRec.AppendEntity(mtext);
                    trans.AddNewlyCreatedDBObject(mtext, true);
                    leader.Annotation = mtext.ObjectId;
                    leader.Annotative = AnnotativeStates.True;
                    leader.EvaluateLeader();
                }
            }
        }
        // Adds a rebar circle with hatching
        private void AddRebarCircle(BlockTableRecord msBlkRec, Transaction trans, Point3d center, double radius, ObjectId layerId)
        {
            // Create rebar circle
            Circle bar = new Circle(center, Vector3d.ZAxis, radius);
            bar.LayerId = layerId;

            // Add the circle to the database FIRST
            msBlkRec.AppendEntity(bar);
            trans.AddNewlyCreatedDBObject(bar, true); // Ensures it's in the DB

            // Create the hatch object
            Hatch hatch = new Hatch();
            hatch.SetDatabaseDefaults();
            hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            hatch.LayerId = layerId; // Assign to the same layer

            // Add hatch to the drawing BEFORE setting reactors
            msBlkRec.AppendEntity(hatch);
            trans.AddNewlyCreatedDBObject(hatch, true);

            // Associate the hatch with the circle (AFTER the circle exists in DB)
            ObjectIdCollection ids = new ObjectIdCollection { bar.ObjectId };
            hatch.AppendLoop(HatchLoopTypes.Default, ids);
            hatch.EvaluateHatch(true);
        }
        private AcadLine TrimLine(AcadLine line, AcadLine otherLine, Point3d? pickPoint = null)
        {
            // Get intersection points
            Point3dCollection intersectionPoints = new Point3dCollection();
            line.IntersectWith(otherLine, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero);

            if (intersectionPoints.Count == 0)
                return null;

            Point3d intersection = intersectionPoints[0]; // Take first intersection

            // Calculate distances from intersection to line endpoints
            double distToStart = intersection.DistanceTo(line.StartPoint);
            double distToEnd = intersection.DistanceTo(line.EndPoint);

            if (pickPoint.HasValue)
            {
                // Use pickPoint to decide which endpoint to keep
                double distPickToStart = pickPoint.Value.DistanceTo(line.StartPoint);
                double distPickToEnd = pickPoint.Value.DistanceTo(line.EndPoint);

                // Keep the endpoint closer to pickPoint
                if (distPickToStart < distPickToEnd)
                    line.EndPoint = intersection; // Trim end point
                else
                    line.StartPoint = intersection; // Trim start point
            }
            else
            {
                // Fallback to original logic: keep segment that doesn't cross otherLine
                Point3d testPoint = distToStart < distToEnd
                    ? line.StartPoint + (line.EndPoint - line.StartPoint).GetNormal() * 0.01
                    : line.EndPoint + (line.StartPoint - line.EndPoint).GetNormal() * 0.01;

                if (!IsPointOnOtherSide(testPoint, otherLine, intersection))
                {
                    if (distToStart > distToEnd)
                        line.EndPoint = intersection;
                    else
                        line.StartPoint = intersection;
                }
                else
                {
                    if (distToStart < distToEnd)
                        line.StartPoint = intersection;
                    else
                        line.EndPoint = intersection;
                }
            }

            return line;
        }

        private AcadLine ExtendLine(AcadLine line, AcadLine otherLine, Point3d? pickPoint = null)
        {
            // Get intersection points
            Point3dCollection intersectionPoints = new Point3dCollection();
            line.IntersectWith(otherLine, Intersect.ExtendThis, intersectionPoints, IntPtr.Zero, IntPtr.Zero);

            if (intersectionPoints.Count == 0)
                return null;

            Point3d intersection = intersectionPoints[0]; // Take first intersection

            // Calculate distances from intersection to line endpoints
            double distToStart = intersection.DistanceTo(line.StartPoint);
            double distToEnd = intersection.DistanceTo(line.EndPoint);

            if (pickPoint.HasValue)
            {
                // Use pickPoint to decide which endpoint to extend
                double distPickToStart = pickPoint.Value.DistanceTo(line.StartPoint);
                double distPickToEnd = pickPoint.Value.DistanceTo(line.EndPoint);

                // Extend the endpoint closer to pickPoint
                if (distPickToStart < distPickToEnd)
                    line.StartPoint = intersection; // Extend start point
                else
                    line.EndPoint = intersection; // Extend end point
            }
            else
            {
                // Fallback to original logic: extend the closer endpoint
                if (distToStart < distToEnd)
                    line.StartPoint = intersection;
                else
                    line.EndPoint = intersection;
            }

            return line;
        }


        private bool IsPointOnOtherSide(Point3d testPoint, AcadLine otherLine, Point3d intersection)
        {
            // Determine if testPoint is on the opposite side of otherLine relative to intersection
            Vector3d lineDir = otherLine.EndPoint - otherLine.StartPoint;
            Vector3d toTest = testPoint - otherLine.StartPoint;
            Vector3d toIntersection = intersection - otherLine.StartPoint;

            double crossTest = lineDir.X * toTest.Y - lineDir.Y * toTest.X;
            double crossIntersection = lineDir.X * toIntersection.Y - lineDir.Y * toIntersection.X;

            return (crossTest * crossIntersection < 0);
        }

        private Point3d? GetIntersectionPoint(AcadLine line1, AcadLine line2)
        {
            Point3dCollection intersectionPoints = new Point3dCollection();
            line1.IntersectWith(line2, Intersect.OnBothOperands, intersectionPoints, IntPtr.Zero, IntPtr.Zero);
            return intersectionPoints.Count > 0 ? intersectionPoints[0] : (Point3d?)null;
        }

        /// <summary>
        /// Creates a line offset from an input line by a specified distance, optionally trimmed by another line.
        /// The offset line is placed on the specified layer or the current drawing layer if none is provided.
        /// </summary>
        /// <param name="msBlkRec">The BlockTableRecord (model space) to add the offset line to.</param>
        /// <param name="trans">The active Transaction for database operations.</param>
        /// <param name="inputLine">The input Line to offset.</param>
        /// <param name="offsetDistance">The distance to offset the line. Positive offsets to the left (relative to the line's direction from start to end), negative to the right.</param>
        /// <param name="trimLine">Optional Line to trim the offset line's endpoints. If null, no trimming is applied.</param>
        /// <param name="layerName">Optional layer name for the offset line. If null, uses the current drawing layer. Defaults to null.</param>
        /// <returns>The created offset Line, added to the model space, or null if creation fails.</returns>
        private AcadLine CreateOffsetLine(BlockTableRecord msBlkRec, Transaction trans, AcadLine inputLine, double offsetDistance, AcadLine trimLine = null, string layerName = null)
        {
            try
            {
                // Get input line's start and end points
                Point3d startPoint = inputLine.StartPoint;
                Point3d endPoint = inputLine.EndPoint;

                // Calculate direction vector of the input line
                Vector3d direction = (endPoint - startPoint).GetNormal();

                // Calculate perpendicular vector (to the left when looking from start to end)
                Vector3d perpVector = new Vector3d(-direction.Y, direction.X, 0).GetNormal();

                // Adjust offset direction based on offsetDistance sign
                // Positive offsetDistance: offset to the left
                // Negative offsetDistance: offset to the right
                Vector3d offsetVector = perpVector * Math.Sign(offsetDistance) * Math.Abs(offsetDistance);

                // Calculate offset line's start and end points
                Point3d offsetStart = startPoint + offsetVector;
                Point3d offsetEnd = endPoint + offsetVector;

                // Create the offset line
                AcadLine offsetLine = new AcadLine(offsetStart, offsetEnd);

                // If trimLine is provided, adjust the offset line's endpoints
                if (trimLine != null)
                {
                    // Ensure both lines are in the database for IntersectWith
                    if (offsetLine.ObjectId.IsNull)
                    {
                        msBlkRec.AppendEntity(offsetLine);
                        trans.AddNewlyCreatedDBObject(offsetLine, true);
                    }
                    if (trimLine.ObjectId.IsNull)
                    {
                        msBlkRec.AppendEntity(trimLine);
                        trans.AddNewlyCreatedDBObject(trimLine, true);
                    }

                    // Find intersections in the XY plane
                    Point3dCollection intersections = new Point3dCollection();
                    offsetLine.IntersectWith(trimLine, Intersect.ExtendBoth, new Plane(), intersections, IntPtr.Zero, IntPtr.Zero);

                    if (intersections.Count > 0)
                    {
                        // Sort intersections by distance from offsetStart
                        List<Point3d> sortedIntersections = new List<Point3d>();
                        foreach (Point3d pt in intersections)
                        {
                            sortedIntersections.Add(pt);
                        }
                        sortedIntersections.Sort((a, b) => a.DistanceTo(offsetStart).CompareTo(b.DistanceTo(offsetStart)));

                        // Update offset line endpoints based on intersections
                        if (sortedIntersections.Count >= 1)
                        {
                            offsetLine.StartPoint = sortedIntersections[0];
                        }
                        if (sortedIntersections.Count >= 2)
                        {
                            offsetLine.EndPoint = sortedIntersections[1];
                        }
                        else
                        {
                            // If only one intersection, trim to that point and keep the other endpoint
                            offsetLine.EndPoint = sortedIntersections[0].DistanceTo(offsetStart) < sortedIntersections[0].DistanceTo(offsetEnd) ? offsetEnd : offsetStart;
                        }
                    }
                    // No intersections: return the full offset line
                }

                // Assign the offset line to the specified or current layer
                ObjectId layerId = layerName != null
                    ? GetOrCreateLayer(msBlkRec.Database, trans, layerName, 4) // Cyan
                    : msBlkRec.Database.Clayer;

                offsetLine.LayerId = layerId;

                // If offsetLine was added temporarily for intersections, ensure it's properly added
                if (offsetLine.ObjectId.IsNull)
                {
                    msBlkRec.AppendEntity(offsetLine);
                    trans.AddNewlyCreatedDBObject(offsetLine, true);
                }

                return offsetLine;
            }
            catch
            {
                return null; // Return null if offset creation fails
            }
        }

            





















        void IExtensionApplication.Initialize() { }
        void IExtensionApplication.Terminate() { }












    }
}
