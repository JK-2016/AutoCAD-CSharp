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
            Editor ed = acadApp.DocumentManager.MdiActiveDocument.Editor;

            // Prompt user for Excel file path
            PromptStringOptions pso = new PromptStringOptions("\nEnter the path to the Excel file: ");
            pso.AllowSpaces = true;
            PromptResult pr = ed.GetString(pso);
            if (pr.Status != PromptStatus.OK) return;
            string excelFilePath = pr.StringResult;

            // Validate file existence
            if (!System.IO.File.Exists(excelFilePath))
            {
                ed.WriteMessage("\nExcel file not found at the specified path.");
                return;
            }

            // Initialize Excel application
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(excelFilePath);

                // Get list of sheet names
                System.Collections.Generic.List<string> sheetNames = new System.Collections.Generic.List<string>();
                foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in workbook.Sheets)
                {
                    sheetNames.Add(sheet.Name);
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
                Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

                // Read dimensions from the second row (assuming first row is headers)
                double lengthWeir = Convert.ToDouble(range.Cells[2, 2].Value) * 1000;
                double crestLevel = Convert.ToDouble(range.Cells[3, 2].Value) * 1000;
                double depthHighCoeff = Convert.ToDouble(range.Cells[4, 2].Value) * 1000;
                double widthStem = Convert.ToDouble(range.Cells[5, 2].Value) * 1000;
                double usBatterWidth = Convert.ToDouble(range.Cells[6, 2].Value) * 1000;
                double dsBatterWidth = Convert.ToDouble(range.Cells[7, 2].Value) * 1000;
                double foundationOffset = Convert.ToDouble(range.Cells[8, 2].Value) * 1000;
                double foundationTopLevel = Convert.ToDouble(range.Cells[9, 2].Value) * 1000;
                double foundationBottomLevel = Convert.ToDouble(range.Cells[10, 2].Value) * 1000;

                // Prompt user to specify insertion point
                PromptPointResult prPtRes1 = ed.GetPoint("\nSpecify Insertion point: ");
                if (prPtRes1.Status != PromptStatus.OK) return;
                Point3d pnt1 = prPtRes1.Value;

                // Calculate additional points
                double crestBaseLevel = crestLevel - depthHighCoeff;
                double baseWidth = widthStem + usBatterWidth + dsBatterWidth;
                double totalWidth = baseWidth + 2 * foundationOffset;
                double xStart = pnt1.X;//- totalWidth / 2;

                Database db = HostApplicationServices.WorkingDatabase;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord msBlkRec = trans.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    ObjectId dimStyleId = GetOrCreateDimStyle(db, trans);

                    // Define points for the weir section
                    Point3d[] Pnts = new Point3d[8];
                    Pnts[0] = new Point3d(xStart, foundationBottomLevel, 0);
                    Pnts[1] = new Point3d(xStart + totalWidth, foundationBottomLevel, 0);

                    Pnts[2] = new Point3d(xStart , foundationTopLevel, 0);
                    Pnts[3] = new Point3d(xStart + totalWidth , foundationTopLevel, 0);


                    Pnts[4] = new Point3d(xStart + foundationOffset, foundationTopLevel, 0);
                    Pnts[5] = new Point3d(xStart + totalWidth - foundationOffset, foundationTopLevel, 0);

                    Pnts[6] = new Point3d(xStart + foundationOffset + usBatterWidth, crestBaseLevel, 0);
               ;

                    Pnts[7] = new Point3d(xStart + foundationOffset + usBatterWidth + widthStem, crestLevel, 0);

                    // Draw the weir outline
                    AddLine(msBlkRec, trans, Pnts[0], Pnts[1], "Wall");
                    AddLine(msBlkRec, trans, Pnts[1], Pnts[3], "Wall");
                    AddLine(msBlkRec, trans, Pnts[3], Pnts[5], "Wall");
                    AddLine(msBlkRec, trans, Pnts[5], Pnts[7], "Wall");
                    AddLine(msBlkRec, trans, Pnts[7], Pnts[6], "Wall");
                    AddLine(msBlkRec, trans, Pnts[6], Pnts[4], "Wall");
                    AddLine(msBlkRec, trans, Pnts[4], Pnts[2], "Wall");
                    AddLine(msBlkRec, trans, Pnts[2], Pnts[0], "Wall");
                  

                    // Add dimensions
                    int dim_dist = 200;
                    AddDimension(msBlkRec, trans, Pnts[0], Pnts[1], new Point3d((Pnts[0].X + Pnts[1].X) / 2, Pnts[0].Y - dim_dist, 0), dimStyleId, 0);
                    AddDimension(msBlkRec, trans, Pnts[2], Pnts[3], new Point3d((Pnts[2].X + Pnts[3].X) / 2, Pnts[2].Y - dim_dist, 0), dimStyleId, 0);
                    AddDimension(msBlkRec, trans, Pnts[4], Pnts[5], new Point3d((Pnts[4].X + Pnts[5].X) / 2, Pnts[4].Y + dim_dist, 0), dimStyleId, 0);
                    AddDimension(msBlkRec, trans, Pnts[6], Pnts[7], new Point3d((Pnts[6].X + Pnts[7].X) / 2, Pnts[6].Y + dim_dist, 0), dimStyleId, 0);
                    AddDimension(msBlkRec, trans, Pnts[0], Pnts[2], new Point3d((Pnts[0].X + Pnts[2].X) / 2, Pnts[0].Y - dim_dist, 0), dimStyleId, 0);
                    AddDimension(msBlkRec, trans, Pnts[3], Pnts[1], new Point3d((Pnts[3].X + Pnts[1].X) / 2, Pnts[1].Y - dim_dist, 0), dimStyleId, 0);

                    AddDimension(msBlkRec, trans, Pnts[0], Pnts[2], new Point3d(Pnts[0].X - dim_dist, (Pnts[0].Y + Pnts[2].Y) / 2, 0), dimStyleId, Math.PI * 0.5);
                    AddDimension(msBlkRec, trans, Pnts[2], Pnts[4], new Point3d(Pnts[2].X - dim_dist, (Pnts[2].Y + Pnts[4].Y) / 2, 0), dimStyleId, Math.PI * 0.5);
                    AddDimension(msBlkRec, trans, Pnts[4], Pnts[6], new Point3d(Pnts[4].X - dim_dist, (Pnts[4].Y + Pnts[6].Y) / 2, 0), dimStyleId, Math.PI * 0.5);

                    trans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError reading Excel file: {ex.Message}");
            }
            finally
            {
                // Clean up Excel objects
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
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
            }
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
        private void AddLine(BlockTableRecord msBlkRec, Transaction trans, Point3d start, Point3d end,String layer)
        {
            Line line = new Line(start, end);
            ObjectId wallLayerId = GetOrCreateLayer(msBlkRec.Database, trans, layer, 1);
            line.LayerId = wallLayerId;
            msBlkRec.AppendEntity(line);
            trans.AddNewlyCreatedDBObject(line, true);
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
        private void AddLeaderWithText(BlockTableRecord msBlkRec, Transaction trans, Point3d attachPoint, string text)
        {
            ObjectId leaderLayerId = GetOrCreateLayer(msBlkRec.Database, trans, "Annotations", 4);
            double dimasz = (double)acadApp.GetSystemVariable("DIMASZ");
            using (Leader leader = new Leader())
            {
                leader.SetDatabaseDefaults();
                leader.LayerId = leaderLayerId;
                leader.Dimasz = 200;
                leader.ColorIndex = 4; // Cyan for annotations
                leader.AppendVertex(attachPoint);
                //leader.AppendVertex(new Point3d(attachPoint.X + 300, attachPoint.Y + 300, 0)); // Leader elbow
                leader.AppendVertex(new Point3d(attachPoint.X + 700, attachPoint.Y + 400, 0)); // Final point
                msBlkRec.AppendEntity(leader);
                trans.AddNewlyCreatedDBObject(leader, true);
                // Attach multiline text to leader
                using (MText mtext = new MText())
                {
                    mtext.SetDatabaseDefaults();
                    mtext.Contents = text;
                    mtext.Location = new Point3d(attachPoint.X + 700, attachPoint.Y + 400, 0);
                    mtext.TextHeight = 150;
                    mtext.LayerId = leaderLayerId;

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

        void IExtensionApplication.Initialize() { }
        void IExtensionApplication.Terminate() { }
    }
}
