using HtmlAgilityPack;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Search;
using MailKit.Security;
using Microsoft.Win32;
using MimeKit;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Octokit;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Catalogs;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Geometry3d;
using ZI_library;
using SD = System.Drawing;
using TSD = Tekla.Structures.Drawing;
using TSDUI = Tekla.Structures.Drawing.UI;
using TSM = Tekla.Structures.Model;

namespace EveryThing
{
    public partial class Form1 : Form
    {

        private void AddEmptyRowsToGridView(int rowstoadd, ref DataGridView dataGridView)
        {
            DataTable dataTable = dataGridView.DataSource as DataTable;

            for (int ii = 0; ii < rowstoadd; ii++)
            {
                DataRow dataRow = dataTable.NewRow();
                dataRow[0] = "";
                dataRow[1] = "";
                dataRow[2] = "";

                dataTable.Rows.Add(dataRow);
            }
        }


        public Form1()
        {
            InitializeComponent();

            //List<string> folders = Tekla.Structures.Dialog.UIControls.EnvironmentFiles.GetStandardPropertyFileDirectories();
            //List<string> multi = Tekla.Structures.Dialog.UIControls.EnvironmentFiles.GetMultiDirectoryFileList("metcon");

            //FileInfo atribFile = Tekla.Structures.Dialog.UIControls.EnvironmentFiles.GetAttributeFile("ConsoleReport.exe.config");
        }

        TSM.Model model;

        private void button1_Click(object sender, EventArgs e)
        {
            TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();
            try
            {
                if (drawingHandler.GetConnectionStatus())
                {
                    TSDUI.Picker picker = drawingHandler.GetPicker();

                    TSD.DrawingObject drawingObject = null;
                    TSD.ViewBase viewBase = null;

                    picker.PickObject("Выберите объект", out drawingObject, out viewBase);

                    if (drawingObject is TSD.Text)
                    {
                        TSD.Text text = drawingObject as TSD.Text;
                        MessageBox.Show(text.Attributes.Font.Color.ToString());
                    }

                    if (drawingObject is TSD.StraightDimensionSet)
                    {
                        TSD.StraightDimension dimension = null;

                        TSD.StraightDimensionSet dimensionSet = drawingObject as TSD.StraightDimensionSet;
                        double totalLength = 0.0;
                        TSD.DrawingObjectEnumerator doe = dimensionSet.GetObjects();
                        if (doe.GetSize() < 2)
                        {
                            MessageBox.Show("Не является цепочкой размеров!", " ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk,
                               MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                            return;
                        }
                        while (doe.MoveNext())
                        {
                            dimension = doe.Current as TSD.StraightDimension;

                            Vector upVector = dimension.UpDirection;

                            Point prjP = Projection.PointToLine(dimension.StartPoint, new Line(dimension.EndPoint, dimension.EndPoint + upVector));
                            double length = Distance.PointToPoint(dimension.StartPoint, prjP);
                            totalLength += length;
                        }
                        MessageBox.Show("Сумма размеров: - " + totalLength, " ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                    }
                    if (drawingObject is TSD.View)
                    {
                        string test = "";
                        drawingObject.Select();
                    }
                }
            }
            catch (TSD.PickerInterruptedException)
            {

            }





            if (drawingHandler.GetConnectionStatus())
            {
                try
                {
                    TSDUI.Picker picker = drawingHandler.GetPicker();

                    TSD.DrawingObject drawingObject = null;
                    TSD.ViewBase viewBase = null;

                    picker.PickObject("Выберите объект", out drawingObject, out viewBase);

                    string guid = "";
                    drawingObject.GetUserProperty(UserPropId, ref guid);

                    if (drawingObject is TSD.StraightDimensionSet)  //*********ДОБАВИЛ В МАКРОСЫ
                    {
                        TSD.StraightDimension dimension = null;

                        TSD.StraightDimensionSet dimensionSet = drawingObject as TSD.StraightDimensionSet;
                        double totalLength = 0.0;
                        TSD.DrawingObjectEnumerator doe = dimensionSet.GetObjects();
                        if (doe.GetSize() < 2)
                        {
                            MessageBox.Show("Не является цепочкой размеров!", " ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk,
                               MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                            return;
                        }
                        while (doe.MoveNext())
                        {
                            dimension = doe.Current as TSD.StraightDimension;

                            Vector upVector = dimension.UpDirection;

                            Point prjP = Projection.PointToLine(dimension.StartPoint, new Line(dimension.EndPoint, dimension.EndPoint + upVector));
                            double length = Distance.PointToPoint(dimension.StartPoint, prjP);
                            totalLength += length;
                        }
                        MessageBox.Show("Сумма размеров: - " + totalLength, " ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                    }




                    if (drawingObject is TSD.Plugin)
                    {
                        TSD.Plugin plugin = drawingObject as TSD.Plugin;
                        plugin.Select();
                        plugin.SetUserProperty("uuu", "test");
                        string test = "";
                        plugin.GetUserProperty("uuu", ref test);
                    }
                    if (drawingObject is TSD.View)
                    {
                        TSD.View view = drawingObject as TSD.View;
                        //view.SetUserProperty("rrr", "UserPropTest");

                        string UPTest = "";
                        view.GetUserProperty("rrr", ref UPTest);
                    }
                    if (drawingObject is TSD.SectionMark)
                    {
                        TSD.SectionMark sectionMark = drawingObject as TSD.SectionMark;
                        sectionMark.Select();
                        string CUST_ID = "";
                        sectionMark.GetUserProperty(CUST_ID, ref CUST_ID);
                        //Script.ChangeSectionMark("bxcv");

                    }
                    if (drawingObject is TSD.DetailMark)
                    {
                        //Script.ChangeDetailMark("det");
                    }
                }
                catch (TSD.PickerInterruptedException)
                {

                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();

                if (drawingHandler.GetConnectionStatus())
                {
                    TSD.ViewBase pickedView = null;
                    TSD.ViewBase sectionMarksView = null;
                    Point startSectionPoint = null;
                    Point endSectionPoint = null;
                    Point sectionViewInsertionPoint = null;
                    TSD.View.ViewAttributes sectionViewAttr = new TSD.View.ViewAttributes("standard");
                    TSD.SectionMarkBase.SectionMarkAttributes sectionMarkAttr = new TSD.SectionMarkBase.SectionMarkAttributes("standard");
                    double depthUp = 1000;
                    double depthDown = 1000;
                    TSD.View sectionView = null;
                    TSD.SectionMark sectionMark = null;

                    TSDUI.Picker picker = drawingHandler.GetPicker();
                    picker.PickTwoPoints("Начало", "Конец", out startSectionPoint, out endSectionPoint, out sectionMarksView);
                    picker.PickPoint("Местоположение вида", out sectionViewInsertionPoint, out pickedView);

                    bool insertSection = TSD.View.CreateSectionView(sectionMarksView as TSD.View, startSectionPoint, endSectionPoint, sectionViewInsertionPoint,
                        depthUp, depthDown, sectionViewAttr, sectionMarkAttr,
                        out sectionView, out sectionMark);

                    //sectionView.Origin = sectionViewInsertionPoint;
                    //sectionView.Modify();

                    //drawingHandler.GetDrawingObjectSelector().UnselectAllObjects();
                    //drawingHandler.GetDrawingObjectSelector().SelectObject(sectionMark);
                    //Script.ChangeSectionMark("12345");
                    //drawingHandler.GetDrawingObjectSelector().UnselectAllObjects();


                }
            }
            catch (Exception ex)
            {
                Messages.TextMessageBox(ex.ToString());
            }
            finally
            {

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();

                TSD.ViewBase pickedView = null;
                TSD.ViewBase sectionMarksView = null;
                Point startSectionPoint = null;
                Point endSectionPoint = null;
                Point sectionViewInsertionPoint = null;
                TSD.View.ViewAttributes sectionViewAttr = new TSD.View.ViewAttributes("standard");
                TSD.SectionMarkBase.SectionMarkAttributes sectionMarkAttr = new TSD.SectionMarkBase.SectionMarkAttributes("standard");
                double depthUp = 1000;
                double depthDown = 1000;
                TSD.View sectionView = null;
                TSD.SectionMark sectionMark = null;

                TSDUI.Picker picker = drawingHandler.GetPicker();
                picker.PickTwoPoints("Начало", "Конец", out startSectionPoint, out endSectionPoint, out sectionMarksView);
                picker.PickPoint("Местоположение вида", out sectionViewInsertionPoint, out pickedView);



                TSD.Drawing Drawing = drawingHandler.GetActiveDrawing();
                CoordinateSystem SectionViewCS = new CoordinateSystem(new Point(0, 0), new Vector(0, 1, 0), new Vector(0, 0, 1));

                sectionView = new TSD.View(Drawing.GetSheet(), SectionViewCS, SectionViewCS,
                    new AABB(new Point(-500, -500, -1000), new Point(500, 10000, 1000)));
                sectionView.Name = "Плагин";
                sectionView.Origin = new Point(50, -50);
                sectionView.Insert();

                SectionViewCS = new CoordinateSystem(new Point(0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                sectionView = new TSD.View(Drawing.GetSheet(), SectionViewCS, SectionViewCS,
                    new AABB(new Point(-500, -500, -1000), new Point(500, 10000, 1000)));
                sectionView.Name = "Ещё плагин";
                sectionView.Origin = new Point(200, -50);

                sectionView.Insert();
            }
            catch (Exception ex)
            {
                Messages.TextMessageBox(ex.ToString());
            }
            finally
            {

            }
        }



        /// <summary>
        /// Задать настройки разреза
        /// </summary>
        static public void ModifyCutSym(string fileName)
        {
            TeklaStructures.Connect();
            MacroBuilder mb = new MacroBuilder();
            mb.CommandEnd();
            mb.Callback("acmd_display_selected_drawing_object_dialog", "", "View_10 window_1");
            mb.ValueChange("csym_dial", "gr_csym_get_menu", fileName);
            mb.PushButton("gr_csym_get", "csym_dial");
            mb.PushButton("csym_modify", "csym_dial");
            mb.PushButton("csym_cancel", "csym_dial");
            mb.Run();

            // ждем пока макрос закончит свою работу
            while (Tekla.Structures.Model.Operations.Operation.IsMacroRunning())
                System.Threading.Thread.Sleep(500);
        }



        public class Script
        {
            private static string xs_macros_directory = "";
            private static string str_begin = "namespace Tekla.Technology.Akit.UserScript\n{public class Script\n{public static void Run(Tekla.Technology.Akit.IScript akit)\n{\n";
            private static string str_end = "\n}\n}\n}";
            public static string _path = "modeling\\";
            public static string _name = "sdr_temp_macros";
            public static string strf = "";


            public static void WriteMacrosAndRun(string str)
            {
                TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref xs_macros_directory);
                string s = "";
                bool tr = false;
                s += xs_macros_directory.ToString();
                tr = s.Contains(";");
                if (tr)
                {

                    string[] s0 = new string[1];
                    s0 = xs_macros_directory.Split(new char[] { ';' });
                    try
                    {
                        for (int i = 0; i <= s0.Length - 1; i++)
                        {
                            if (s0[i] != null || s0[i] != "")
                            {
                                if (@s0[i].Substring(s0[i].Length - 1, 1) != "\\")
                                    @s0[i] += "\\";
                                if (write((@s0[i] + @_path + _name), (@str_begin + @str + @str_end)))
                                {
                                    i = s0.Length;
                                    try
                                    {
                                        File.Delete(strf + ".cs");
                                        File.Delete(strf + ".dll");
                                        File.Delete(strf + ".pdb");
                                    }
                                    catch { }
                                }
                                else
                                {
                                    try
                                    {
                                        File.Delete(strf + ".cs");
                                        File.Delete(strf + ".dll");
                                        File.Delete(strf + ".pdb");
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error!" + "\n" + ex.ToString());
                    }
                }
                else
                {
                    if (@xs_macros_directory.Substring(@xs_macros_directory.Length - 1, 1) != "\\")
                        @xs_macros_directory += "\\";
                    if (!write((@xs_macros_directory + @_path + @_name), (@str_begin + @str + @str_end))) MessageBox.Show("Error");
                    try
                    {
                        File.Delete(strf + ".cs");
                        File.Delete(strf + ".dll");
                        File.Delete(strf + ".pdb");
                    }
                    catch { }
                }

            }
            public static bool write(string @str, string _string)
            {
                StreamWriter _file = File.CreateText(str + ".cs");
                _file.WriteLine(_string);
                _file.Close();
                // Run Macro
                TSM.Model model = new TSM.Model();
                if (model.GetConnectionStatus())
                {
                    string ss = str.Replace(@"\", @"\\");
                    strf = ss;
                    return (Tekla.Structures.Model.Operations.Operation.RunMacro(_name + ".cs"));
                }
                else return false;
                // return false;
            }

            public static void ChangeSectionMark(string t)
            {
                string str =
                @"akit.Callback(""acmd_display_attr_dialog"", ""csym_dial"", ""main_frame"");" +
                @"akit.PushButton(""csym_on_off"", ""csym_dial"");" +
                @"akit.ValueChange(""csym_dial"", ""csym_label"",""" + t + @"""" + ");";
                str += @"akit.PushButton(""csym_modify"", ""csym_dial"");";
                str += @"akit.PushButton(""csym_cancel"", ""csym_dial"");";
                WriteMacrosAndRun(str);
            }
            public static void ChangeDetailMark(string t)
            {
                string str =
                @"akit.Callback(""acmd_display_attr_dialog"", ""detail_dial"", ""main_frame"");" +
                @"akit.ValueChange(""detail_dial"", ""lbltxtDetailLabelIndexStart"",""" + t + @"""" + ");";
                str += @"akit.PushButton(""butDetailSymbol_modify"", ""detail_dial"");";
                str += @"akit.PushButton(""butDetailSymbol_cancel"", ""detail_dial"");";
                WriteMacrosAndRun(str);
            }
        }






        private void CreateFilterFile(TSM.Model model, string partClass) //Создаем файл для фильтра выбора
        {
            string sysPath = "";
            sysPath = model.GetInfo().ModelPath;
            sysPath = Path.Combine(sysPath, "attributes");
            sysPath = Path.Combine(sysPath, "Ficep.SObjGrp");
            StreamWriter FF = new StreamWriter(sysPath, false, Encoding.Default);
            FF.WriteLine("TITLE_OBJECT_GROUP");
            FF.WriteLine("{");
            FF.WriteLine("    Version= 1.05");
            FF.WriteLine("    Count= 1");
            FF.WriteLine("    SECTION_OBJECT_GROUP");
            FF.WriteLine("    {");
            FF.WriteLine("        0");
            FF.WriteLine("        1");
            FF.WriteLine("        co_part");
            FF.WriteLine("        proCLASS");
            FF.WriteLine("        albl_Class");
            FF.WriteLine("        ==");
            FF.WriteLine("        albl_Equals");
            FF.WriteLine("        " + partClass);
            FF.WriteLine("        0");
            FF.WriteLine("        Empty");
            FF.WriteLine("        }");
            FF.WriteLine("}");
            FF.Close();
        }



        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                TSM.Model model = new TSM.Model();

                if (model.GetConnectionStatus())
                {
                    Tekla.Structures.Model.UI.ModelObjectSelector selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    TSM.ModelObjectEnumerator SelectBlock = selector.GetSelectedObjects();
                    SelectBlock.SelectInstances = false;

                    if (SelectBlock.GetSize() > 0)
                    {
                        while (SelectBlock.MoveNext())
                        {
                            TSM.ModelObject modelObject = SelectBlock.Current as TSM.ModelObject;
                            //modelObject.Select();

                            modelObject.Select();
                            modelObject.SetUserProperty("ES_DESC", 5);

                            if (modelObject is TSM.Beam)
                            {
                                TSM.Beam beam = modelObject as TSM.Beam;
                                beam.Select();
                                string partClass = beam.Class;
                                string partProfile = beam.Profile.ProfileString;

                                CreateFilterFile(model, partClass);
                            }


                            if (ModifierKeys == Keys.Shift)
                            {

                            }


                            Hashtable userPropTable = new Hashtable();
                            modelObject.GetAllUserProperties(ref userPropTable);

                            //if (userPropTable.Count == 0)
                            //    return;

                            //modelObject.SetUserProperty("metcon_WBNumber", 68);

                            modelObject.SetUserProperty("ACTUAL_FAB_DATE", 1546732800);

                            IDictionaryEnumerator properties = userPropTable.GetEnumerator();
                            while (properties.MoveNext())
                            {
                                modelObject.SetUserProperty(properties.Key.ToString(), 88);
                            }
                            userPropTable = new Hashtable();
                            modelObject.GetAllUserProperties(ref userPropTable);
                            modelObject.Select();
                            //modelObject.Delete();

                            model.CommitChanges();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                model = new TSM.Model();

                if (model.GetConnectionStatus())
                {
                    Tekla.Structures.Model.UI.ModelObjectSelector selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    TSM.ModelObjectEnumerator SelectBlock = selector.GetSelectedObjects();
                    SelectBlock.SelectInstances = false;

                    if (SelectBlock.GetSize() > 0)
                    {
                        while (SelectBlock.MoveNext())
                        {
                            TSM.ModelObject modelObject = SelectBlock.Current as TSM.ModelObject;
                            modelObject.Select();

                            string owner = "";
                            modelObject.GetReportProperty("OWNER", ref owner);

                            if (modelObject is TSM.BoltGroup)
                            {
                                //BoltGroup boltGroup = modelObject as BoltGroup;

                                //bool correct = BoltUtils.CorrectBoltHeadSide(model, boltGroup);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        ArrayList dataList;


        private void button4_Click(object sender, EventArgs e)
        {
            string def = "";
            TeklaStructuresSettings.GetAdvancedOption("XS_​APPLICATIONS_​PATH", ref def);






            CatalogHandler catalogHandler = new CatalogHandler();
            dataList = new ArrayList();

            ProfileItemEnumerator profileItemEnumerator = catalogHandler.GetProfileItems();

            while (profileItemEnumerator.MoveNext())
            {
                if (!(profileItemEnumerator.Current is ParametricProfileItem))
                {
                    LibraryProfileItem PI = profileItemEnumerator.Current as LibraryProfileItem;

                    ArrayList AnalysisParametersAL = PI.aProfileItemAnalysisParameters;
                    List<ProfileItemParameter> AnalysisParametersL = AnalysisParametersAL.Cast<ProfileItemParameter>().ToList();

                    ArrayList UserParametersAL = PI.aProfileItemUserParameters;
                    List<ProfileItemParameter> UserParametersL = UserParametersAL.Cast<ProfileItemParameter>().ToList();

                    ArrayList ProfParamsAL = PI.aProfileItemParameters;
                    List<ProfileItemParameter> ProfParamsL = ProfParamsAL.Cast<ProfileItemParameter>().ToList();

                    string NameCat = PI.ProfileName;
                    string GOSTCat = "";
                    double AreaCat = 0.0;
                    double WeightPerMeterCat = 0.0;
                    double FlangeThick = 0.0;
                    double FlangeThick1 = 0.0;
                    double FlangeThick2 = 0.0;
                    double PlateThick = 0.0;

                    if (AnalysisParametersL.Count > 0)
                    {
                        AreaCat = AnalysisParametersL[1].Value;
                        WeightPerMeterCat = AnalysisParametersL[24].Value;
                    }

                    if (UserParametersL.Count > 0)
                    {
                        GOSTCat = UserParametersL[0].StringValue;
                    }

                    if (ProfParamsL.Count > 0)
                    {
                        for (int i = 0; i < ProfParamsL.Count; i++)
                        {
                            if (ProfParamsL[i].Property == "FLANGE_THICKNESS")
                                FlangeThick = ProfParamsL[i].Value;
                            if (ProfParamsL[i].Property == "FLANGE_THICKNESS_1")
                                FlangeThick1 = ProfParamsL[i].Value;
                            if (ProfParamsL[i].Property == "FLANGE_THICKNESS_2")
                                FlangeThick2 = ProfParamsL[i].Value;
                            if (ProfParamsL[i].Property == "PLATE_THICKNESS")
                                PlateThick = ProfParamsL[i].Value;
                        }
                    }

                    var data = new
                    {
                        ИмяПрофиля = NameCat,
                        ГОСТ = GOSTCat,
                        ПлощадьСечения = AreaCat,
                        МассаПМ = WeightPerMeterCat,
                        ТолщинаПолки = FlangeThick,
                        ТолщинаПолки1 = FlangeThick1,
                        ТолщинаПолки2 = FlangeThick2,
                        ТолщинаПластины = PlateThick
                    };
                    dataList.Add(data);
                }
            }

            DirectoryInfo exportProfileDirectory = new DirectoryInfo(@"Z:\ExportProfileStandard\");
            if (!exportProfileDirectory.Exists)
                exportProfileDirectory.Create();

            string filenameJSON = "Profiles.json";
            JsonSerializer serializer = new JsonSerializer();

            string filePath = Path.Combine(exportProfileDirectory.FullName, filenameJSON);

            using (StreamWriter file = File.CreateText(filePath))
            {
                serializer.Serialize(file, dataList);
            }

            Messages.TextMessageBox("Файл сохранён: " + filePath);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                model = new TSM.Model();

                if (model.GetConnectionStatus())
                {
                    Tekla.Structures.Model.UI.ModelObjectSelector selector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    TSM.ModelObjectEnumerator SelectBlock = selector.GetSelectedObjects();
                    SelectBlock.SelectInstances = false;

                    if (SelectBlock.GetSize() > 0)
                    {
                        while (SelectBlock.MoveNext())
                        {
                            if (SelectBlock.Current is TSM.BoltGroup)
                            {
                                TSM.BoltGroup boltGroup = SelectBlock.Current as TSM.BoltGroup;
                                boltGroup.Select();

                                CoordinateSystem coordinateSystem = boltGroup.GetCoordinateSystem();
                                //cs_net_lib.UI.DrawPlane(coordinateSystem);

                                ArrayList positionsAL = boltGroup.BoltPositions;
                                List<Point> positions = positionsAL.Cast<Point>().ToList();
                                Point firstBoltPoint = positions[0];
                                //GraphDrawer.graphicsDrawer.DrawText(firstBoltPoint, "firstBoltPoint", GraphDrawer.black);

                                Vector zVector = coordinateSystem.AxisX.Cross(coordinateSystem.AxisY);

                                double normalize = 1000.0;

                                zVector.Normalize(1000);
                                Point ls1 = new Point(firstBoltPoint.X, firstBoltPoint.Y, firstBoltPoint.Z);
                                ls1.Translate(zVector.X, zVector.Y, zVector.Z);
                                zVector.Normalize(-normalize);
                                Point ls2 = new Point(firstBoltPoint.X, firstBoltPoint.Y, firstBoltPoint.Z);
                                ls2.Translate(zVector.X, zVector.Y, zVector.Z);

                                LineSegment lineSegment = new LineSegment(ls1, ls2);
                                GraphDrawer.graphicsDrawer.DrawText(ls1, "ls1", GraphDrawer.black);
                                GraphDrawer.graphicsDrawer.DrawText(ls2, "ls2", GraphDrawer.black);

                                ArrayList intersectPointsAL = boltGroup.GetSolid(true).Intersect(lineSegment);
                                List<Point> intersectPoints = intersectPointsAL.Cast<Point>().ToList();
                                for (int i = 0; i < intersectPoints.Count; i++)
                                {
                                    GraphDrawer.graphicsDrawer.DrawText(intersectPoints[i], "Т." + i.ToString(), GraphDrawer.black);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            TSD.DrawingHandler DrawingHandler = new TSD.DrawingHandler();
            int counter = 0;

            if (DrawingHandler.GetConnectionStatus())
            {
                TSD.Drawing CurrentDrawing = DrawingHandler.GetActiveDrawing();
                if (CurrentDrawing != null)
                {
                    TSD.DrawingObjectEnumerator markList = CurrentDrawing.GetSheet().GetAllObjects(typeof(TSD.Mark));
                    while (markList.MoveNext())
                    {
                        TSD.Mark currentMark = markList.Current as TSD.Mark;

                        if (currentMark != null)
                        {
                            IEnumerable enumerable = currentMark.Attributes.Content;
                            foreach (var item in enumerable)
                            {
                                if (item is TSD.TextElement)
                                {
                                    TSD.TextElement text = item as TSD.TextElement;
                                    if (text.Value == "t")
                                    {
                                        text.Value = "-";
                                        counter++;
                                    }
                                }
                            }
                            currentMark.Modify();
                        }
                    }
                    ZI_library.Messages.TextMessageBox("Заменено символов: " + counter.ToString());
                    this.Enabled = true;
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SqlConnection Connection = new SqlConnection(@"Data Source=ZZMK-SRV1; Initial Catalog=weld; Persist Security Info=True; User ID = konstr; Password=konstr");
            if (Connection.State != ConnectionState.Open)
                Connection.Open();

            DataTable weldTables = new DataTable("WeldTables");
            string sql = @"SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' ORDER BY TABLE_NAME";
            SqlCommand command = new SqlCommand(sql, Connection);
            SqlDataAdapter da = new SqlDataAdapter(command);
            da.Fill(weldTables);

            for (int i = 0; i < weldTables.Rows.Count; i++)
            {
                string tableName = weldTables.Rows[i]["TABLE_NAME"].ToString();

                // Добавляем во все таблицы столбик Opt и вставляем в него значение 1
                if (tableName.Contains("weldInAssembly") | tableName.Contains("weld"))
                {
                    //sql = "IF COL_LENGTH('"+ tableName + "', 'Opt') IS NULL ALTER TABLE " + tableName + " ADD Opt VARCHAR(2) NULL";
                    //command = new SqlCommand(sql, Connection);
                    //command.ExecuteNonQuery();

                    //sql = "UPDATE " + tableName + " SET Opt='1'";
                    //command = new SqlCommand(sql, Connection);
                    //command.ExecuteNonQuery();
                }
            }

            ZI_library.Messages.MessageBoxDone();
        }














        private const string ColumnNameDrawingMark = "Метка черт.";
        private const string ColumnNameDrawingName = "Имя черт.";
        private const string ColumnNameSecName = "Имя разреза";
        private const string ColumnNameDetName = "Имя узла";
        private const string ColumnNameSheet = "На листе";
        private const string ColumnNameSecNameChild = "Родит. разрез";
        private const string ColumnNameDetNameChild = "Родит. узел";
        private const string ColumnNameID = "ID";
        private const string UserPropId = "CUST_ID";
        private const string SectionDGVTag = "SectionDGV";
        private const string DetailDGVTag = "DetailDGV";

        private Dictionary<string, TSD.Drawing> nameDraw;


        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                this.Enabled = false;
                string buttonText = button_SecDetailEnum.Text;


                int startSection = -1;
                if (!Int32.TryParse(textBox_SecStart.Text, out startSection))
                {
                    ZI_library.Messages.TextMessageBox("Не удаётся преобразовать начальный номер разреза в целое число!");
                    return;
                }
                int startDetail = -1;
                if (!Int32.TryParse(textBox_DetailStart.Text, out startDetail))
                {
                    ZI_library.Messages.TextMessageBox("Не удаётся преобразовать начальный номер узла в целое число!");
                    return;
                }

                TSD.DrawingHandler dh = new TSD.DrawingHandler();

                // Проходим по выбранным чертежам G
                sections = new DataTable("SectionsDT");
                details = new DataTable("DetailsDT");
                nameDraw = new Dictionary<string, TSD.Drawing>();
                dataGridView_Sec.DataSource = null;
                dataGridView_Det.DataSource = null;
                int drawCounter = 0;
                TSD.DrawingEnumerator GADraw = dh.GetDrawingSelector().GetSelected();
                GADraw.SelectInstances = false;
                while (GADraw.MoveNext())
                {
                    TSD.Drawing drawing = GADraw.Current as TSD.Drawing;
                    nameDraw.Add(drawing.Mark, drawing);
                    if (drawing is TSD.GADrawing)
                    {
                        // c закрытым чертежом не работает SetUserProperty
                        dh.SetActiveDrawing(drawing, false);
                        drawCounter++;
                        button_SecDetailEnum.Text = drawCounter.ToString() + " из " + GADraw.GetSize().ToString();
                        button_SecDetailEnum.Refresh();

                        Type[] Types = new Type[2];
                        Types.SetValue(typeof(TSD.SectionMark), 0);
                        Types.SetValue(typeof(TSD.DetailMark), 1);
                        TSD.DrawingObjectEnumerator sectionsDetails = drawing.GetSheet().GetAllObjects(Types);
                        foreach (TSD.DrawingObject drawingObject in sectionsDetails)
                        {
                            string custID = "";
                            if (drawingObject is TSD.SectionMark)
                            {
                                SetDrawObjId(drawingObject, ref custID);
                                AddRowInDT(drawing, sections, drawingObject, custID);
                            }
                            if (drawingObject is TSD.DetailMark)
                            {
                                SetDrawObjId(drawingObject, ref custID);
                                AddRowInDT(drawing, details, drawingObject, custID);
                            }
                        }
                        dh.CloseActiveDrawing(true);
                    }
                }
                dataGridView_Sec.DataSource = sections;
                dataGridView_Det.DataSource = details;

                if (dataGridView_Sec.Rows.Count > 0)
                    SetPropertiesDGV(dataGridView_Sec);
                if (dataGridView_Det.Rows.Count > 0)
                    SetPropertiesDGV(dataGridView_Det);

                contextMenuStrip = new ContextMenuStrip();
                contextMenuStrip.Items.Add("Привязать");
                contextMenuStrip.Items.Add("Открыть чертёж");

                //TSD.DrawingEnumerator SelectedDrawings = dh.GetDrawingSelector().GetSelected();
                //SelectedDrawings.SelectInstances = false;

                //while (SelectedDrawings.MoveNext())
                //{
                //    TSD.Drawing drawing = SelectedDrawings.Current as TSD.Drawing;
                //    TSD.DrawingUpToDateStatus upToDateStatus = drawing.UpToDateStatus;
                //    if (upToDateStatus.ToString().ToLower().Contains("delete"))
                //        continue;

                //    dh.SetActiveDrawing(drawing, true);

                //    Type[] Types = new Type[2];
                //    Types.SetValue(typeof(TSD.SectionMark), 0);
                //    Types.SetValue(typeof(TSD.DetailMark), 1);
                //    TSD.DrawingObjectEnumerator sectionsDetails = drawing.GetSheet().GetAllObjects(Types);

                //    foreach (TSD.DrawingObject drawingObject in sectionsDetails)
                //    {
                //        if (checkBox_SecEnum.Checked & drawingObject is TSD.SectionMark)
                //        {
                //            dh.GetDrawingObjectSelector().SelectObject(drawingObject);
                //            Script.ChangeSectionMark(startSection.ToString());
                //            startSection++;
                //        }
                //        if (checkBox_DetailEnum.Checked & drawingObject is TSD.DetailMark)
                //        {
                //            dh.GetDrawingObjectSelector().SelectObject(drawingObject);
                //            Script.ChangeDetailMark(startDetail.ToString());
                //            startDetail++;
                //        }
                //    }
                //    dh.CloseActiveDrawing(true);
                //}
                ////// Раскоментировать в плагине
                ////ZI_library.MyControls.TextOnStatusStrip(statusStrip1, "Ожидание...");
                //textBox_SecStart.Text = startSection.ToString();
                //textBox_DetailStart.Text = startDetail.ToString();
                //ZI_library.Messages.MessageBoxDone();
            }
            catch (Exception ex)
            {
                ZI_library.Messages.TextMessageBox(ex.ToString());
            }
            finally
            {
                button_SecDetailEnum.Text = "Пронумеровать";
                this.Enabled = true;
            }
        }

        private void AddRowInDT(TSD.Drawing drawing, DataTable dataTable, TSD.DrawingObject drawingObject, string moID)
        {
            DataRow dataRow = dataTable.NewRow();

            DataColumn DrawingMark = new DataColumn(ColumnNameDrawingMark);
            if (!dataTable.Columns.Contains(DrawingMark.ColumnName))
                dataTable.Columns.Add(DrawingMark);
            dataRow[ColumnNameDrawingMark] = drawing.Mark;

            DataColumn DrawingName = new DataColumn(ColumnNameDrawingName);
            if (!dataTable.Columns.Contains(DrawingName.ColumnName))
                dataTable.Columns.Add(DrawingName);
            dataRow[ColumnNameDrawingName] = drawing.Name;

            DataColumn MarkName = new DataColumn();
            if (drawingObject is TSD.SectionMark)
                MarkName.ColumnName = ColumnNameSecName;
            if (drawingObject is TSD.DetailMark)
                MarkName.ColumnName = ColumnNameDetName;
            if (!dataTable.Columns.Contains(MarkName.ColumnName))
                dataTable.Columns.Add(MarkName);

            DataColumn OnSheet = new DataColumn();
            OnSheet.ColumnName = ColumnNameSheet;
            if (!dataTable.Columns.Contains(OnSheet.ColumnName))
                dataTable.Columns.Add(OnSheet);

            DataColumn MarkNameChild = new DataColumn();
            if (drawingObject is TSD.SectionMark)
                MarkNameChild.ColumnName = ColumnNameSecNameChild;
            if (drawingObject is TSD.DetailMark)
                MarkNameChild.ColumnName = ColumnNameDetNameChild;
            if (!dataTable.Columns.Contains(MarkNameChild.ColumnName))
                dataTable.Columns.Add(MarkNameChild);

            DataColumn ID = new DataColumn();
            ID.ColumnName = ColumnNameID;
            if (!dataTable.Columns.Contains(ID.ColumnName))
                dataTable.Columns.Add(ID);

            if (drawingObject is TSD.SectionMark)
            {
                TSD.SectionMark sectionMark = drawingObject as TSD.SectionMark;
                sectionMark.Select();
                dataRow[ColumnNameSecName] = sectionMark.Attributes.MarkName;
                dataRow[ColumnNameID] = moID;

                TSD.SectionMarkBase.SectionMarkTagsAttributes tags = sectionMark.Attributes.TagsAttributes;
                TSD.ContainerElement container = tags.TagA1.TagContent;
                IEnumerator enumerator = container.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current is TSD.PropertyElement)
                    {
                        TSD.PropertyElement prop = enumerator.Current as TSD.PropertyElement;
                        if (prop.Name == "SOURCE_DRAWING_NAME_WHEN_MOVED")
                        {
                            dataRow[ColumnNameSheet] = prop.Value;
                        }
                    }
                }
            }
            if (drawingObject is TSD.DetailMark)
            {
                TSD.DetailMark detail = drawingObject as TSD.DetailMark;
                detail.Select();
                dataRow[ColumnNameDetName] = detail.Attributes.MarkName;
                dataRow[ColumnNameID] = moID;

                TSD.DetailMarkTagsAttributes tags = detail.Attributes.TagsAttributes;
                TSD.ContainerElement container = tags.TagA2.TagContent;
                IEnumerator enumerator = container.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current is TSD.PropertyElement)
                    {
                        TSD.PropertyElement prop = enumerator.Current as TSD.PropertyElement;
                        if (prop.Name == "SOURCE_DRAWING_NAME")
                        {
                            dataRow[ColumnNameSheet] = prop.Value;
                        }
                    }
                }
            }
            dataTable.Rows.Add(dataRow);
        }

        private DataGridView SetPropertiesDGV(DataGridView dataGridView)
        {
            int columnWidth = 43;
            dataGridView.Columns[ColumnNameDrawingMark].Width = columnWidth;
            dataGridView.Columns[ColumnNameDrawingName].Width = columnWidth;
            if (dataGridView.Columns.Contains(ColumnNameSecName))
                dataGridView.Columns[ColumnNameSecName].Width = columnWidth;
            if (dataGridView.Columns.Contains(ColumnNameDetName))
                dataGridView.Columns[ColumnNameDetName].Width = columnWidth;
            if (dataGridView.Columns.Contains(ColumnNameSecNameChild))
                dataGridView.Columns[ColumnNameSecNameChild].Width = columnWidth;
            if (dataGridView.Columns.Contains(ColumnNameSheet))
                dataGridView.Columns[ColumnNameSheet].Width = columnWidth;
            if (dataGridView.Columns.Contains(ColumnNameDetNameChild))
                dataGridView.Columns[ColumnNameDetNameChild].Width = columnWidth;

            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.AllowUserToResizeRows = false;

            dataGridView_Sec.ReadOnly = true;
            dataGridView_Det.ReadOnly = true;

            return dataGridView;
        }
        DataTable sections;
        DataTable details;
        ContextMenuStrip contextMenuStrip;


        /// <summary>
        /// Создаём айдишники, чтобы можно было в дальнейшем идентифицировать разрез/узел
        /// </summary>
        private void SetDrawObjId(TSD.DrawingObject drawingObject, ref string custID)
        {
            string guid = Guid.NewGuid().ToString();

            drawingObject.Select();
            string custId = "";
            drawingObject.GetUserProperty(UserPropId, ref custId);
            if (custId == "")
                drawingObject.SetUserProperty(UserPropId, guid);
            custID = custId;
        }

        int selectedRowDGV;
        DataGridView selectedDGV;

        private void dataGridViewRightMouseClick(object sender, MouseEventArgs e)
        {
            selectedDGV = sender as DataGridView;
            if (e.Button == MouseButtons.Right)
            {
                ContextMenuStrip menu = new ContextMenuStrip();
                if (selectedDGV.Tag.ToString() == "SectionDGV")
                    menu.Items.Add("Привязать разрез").Name = "Привязать";
                if (selectedDGV.Tag.ToString() == "DetailDGV")
                    menu.Items.Add("Привязать узел").Name = "Привязать";
                menu.Items.Add("Открыть чертёж").Name = "Открыть чертёж";

                selectedRowDGV = selectedDGV.HitTest(e.X, e.Y).RowIndex;

                //if (currentMouseOverRow >= 0)
                //{
                //    menu.MenuItems.Add(new MenuItem(string.Format("Do something to row {0}", currentMouseOverRow.ToString())));
                //}

                menu.Show(selectedDGV, new SD.Point(e.X, e.Y));
                menu.ItemClicked += ContextMenuDGV_ItemClicked;
            }
        }

        private void ContextMenuDGV_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();
            TSD.Drawing drawing = null;
            nameDraw.TryGetValue(selectedDGV[ColumnNameDrawingMark, selectedRowDGV].Value.ToString(), out drawing);

            string ПривязатьColumnName = "";
            string pickerMsg = "";
            if (selectedDGV.Tag.ToString() == SectionDGVTag)
            {
                ПривязатьColumnName = ColumnNameSecNameChild;
                pickerMsg = "Выберите разрез";
            }
            if (selectedDGV.Tag.ToString() == DetailDGVTag)
            {
                ПривязатьColumnName = ColumnNameDetNameChild;
                pickerMsg = "Выберите узел";
            }

            switch (e.ClickedItem.Name.ToString())
            {
                case "Привязать":
                    if (drawingHandler.GetActiveDrawing() != null)
                    {
                        TSDUI.Picker picker = drawingHandler.GetPicker();
                        TSD.DrawingObject drawingObject = null;
                        TSD.ViewBase viewBase = null;
                        picker.PickObject(pickerMsg, out drawingObject, out viewBase);

                        if (selectedDGV.Tag.ToString() == SectionDGVTag)
                        {
                            if (drawingObject is TSD.SectionMark)
                            {
                                TSD.SectionMark sectionMark = drawingObject as TSD.SectionMark;

                                BindingSource src = new BindingSource();
                                src.DataSource = selectedDGV.DataSource;
                                int pos = src.Position = src.Find(ColumnNameSecName, sectionMark.Attributes.MarkName.ToString());

                                selectedDGV[ColumnNameSecName, pos].Value = selectedDGV[ColumnNameSecName, selectedRowDGV].Value.ToString();


                                //selectedDGV[ColumnNameSecName, index].Value= selectedDGV[ColumnNameSecName, selectedRowDGV].Value.ToString();

                                //selectedDGV[ColumnNameSecName, selectedRowDGV].Value=
                                sectionMark.Attributes.MarkName = selectedDGV[ColumnNameSecName, selectedRowDGV].Value.ToString();
                                sectionMark.Modify();
                                //selectedDGV[ПривязатьColumnName, selectedRowDGV].Value = selectedDGV[ColumnNameSecName, selectedRowDGV].Value.ToString();
                            }
                            else
                            {
                                ZI_library.Messages.TextMessageBox("Выбранный объект не является разрезом!");
                                return;
                            }
                        }
                        if (selectedDGV.Tag.ToString() == DetailDGVTag)
                        {
                            if (drawingObject is TSD.DetailMark)
                            {
                                selectedDGV[ПривязатьColumnName, selectedRowDGV].Value = "333";
                            }
                            else
                            {
                                ZI_library.Messages.TextMessageBox("Выбранный объект не является узлом!");
                                return;
                            }
                        }
                    }
                    else
                    {
                        ZI_library.Messages.TextMessageBox("Не открыт чертёж!");
                        return;
                    }
                    break;
                case "Открыть чертёж":
                    if (!drawingHandler.SetActiveDrawing(drawing, true))
                        ZI_library.Messages.TextMessageBox("Не удалось открыть чертёж!");
                    break;
                default:
                    break;
            }
        }












        public class Example
        {
            TSD.Drawing MyCurrentDrawing = new TSD.GADrawing();
            TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();

            ArrayList GetDrawingObjectsByType(Type objectType)
            {
                ArrayList ObjectsToBeSelected = new ArrayList();

                foreach (TSD.DrawingObject drawingObject in MyCurrentDrawing.GetSheet().GetAllObjects())
                {
                    if (drawingObject.GetType() == objectType)
                        ObjectsToBeSelected.Add(drawingObject);
                }

                return ObjectsToBeSelected;
            }

            public void selectSections()
            {
                ArrayList sections = GetDrawingObjectsByType(typeof(TSD.SectionMark));
                if (sections.Count > 0)
                    drawingHandler.GetDrawingObjectSelector().SelectObjects(sections, false);
            }

            public void selectDetails()
            {
                ArrayList details = GetDrawingObjectsByType(typeof(TSD.DetailMark));
                if (details.Count > 0)
                    drawingHandler.GetDrawingObjectSelector().SelectObjects(details, false);
            }
        }



        private void button11_Click(object sender, EventArgs e)
        {
            SD.Color myRgbColor = new SD.Color();
            myRgbColor = SD.Color.Black;
            string text = @"
1 Технические требования см. лист 1.1
2 Маркировку наносить как на стенке, так и на полках
клеймением и маркиратором на расстоянии 500 мм от края
Пример маркировки: 109-91SR301-1-B130
";
            DrawText(text, new SD.Font("GOST type A", 16), myRgbColor, 500, "c:\\1\\1111.png");





        }

        /// <summary>
        /// Converting text to image (png).
        /// </summary>
        /// <param name="text">text to convert</param>
        /// <param name="font">Font to use</param>
        /// <param name="textColor">text color</param>
        /// <param name="maxWidth">max width of the image</param>
        /// <param name="path">path to save the image</param>
        public static void DrawText(String text, SD.Font font, SD.Color textColor, int maxWidth, String path)
        {
            //first, create a dummy bitmap just to get a graphics object
            SD.Bitmap img = new SD.Bitmap(1, 1);
            float dpix = float.Parse("300");
            float dpiy = float.Parse("300");
            img.SetResolution(dpix, dpiy);

            SD.Graphics drawing = SD.Graphics.FromImage(img);
            //measure the string to see how big the image needs to be
            SD.SizeF textSize = drawing.MeasureString(text, font, maxWidth);

            //set the stringformat flags to rtl
            SD.StringFormat sf = new SD.StringFormat();
            //uncomment the next line for right to left languages
            //sf.FormatFlags = StringFormatFlags.DirectionRightToLeft;
            sf.Trimming = SD.StringTrimming.Word;
            //free up the dummy image and old graphics object
            img.Dispose();
            drawing.Dispose();

            //create a new image of the right size
            img = new SD.Bitmap((int)textSize.Width, (int)textSize.Height);

            drawing = SD.Graphics.FromImage(img);
            //Adjust for high quality
            drawing.CompositingQuality = SD.Drawing2D.CompositingQuality.HighQuality;
            drawing.InterpolationMode = SD.Drawing2D.InterpolationMode.HighQualityBilinear;
            drawing.PixelOffsetMode = SD.Drawing2D.PixelOffsetMode.HighQuality;
            drawing.SmoothingMode = SD.Drawing2D.SmoothingMode.HighQuality;
            drawing.TextRenderingHint = SD.Text.TextRenderingHint.SingleBitPerPixel;

            //paint the background
            drawing.Clear(SD.Color.White);

            //create a brush for the text
            SD.Brush textBrush = new SD.SolidBrush(textColor);

            drawing.DrawString(text, font, textBrush, new SD.RectangleF(0, 0, textSize.Width, textSize.Height), sf);

            drawing.Save();

            textBrush.Dispose();
            drawing.Dispose();
            img.Save(path, SD.Imaging.ImageFormat.Png);
            img.Dispose();

        }

        private void button12_Click(object sender, EventArgs e)
        {
            string text = @"1 Технические требования см. лист 1.1
                    2 Маркировку наносить как на стенке, так и на полках
                    клеймением и маркиратором на расстоянии 500 мм от края
                    Пример маркировки: 109-91SR301-1-B130
                    ";
            SD.Bitmap bitmap = new SD.Bitmap(1, 1);
            SD.Font font = new SD.Font("Arial", 25, SD.FontStyle.Regular, SD.GraphicsUnit.Pixel);
            SD.Graphics graphics = SD.Graphics.FromImage(bitmap);
            int width = (int)graphics.MeasureString(text, font).Width;
            int height = (int)graphics.MeasureString(text, font).Height;
            bitmap = new SD.Bitmap(bitmap, new SD.Size(width, height));
            graphics = SD.Graphics.FromImage(bitmap);
            graphics.Clear(SD.Color.White);
            graphics.SmoothingMode = SD.Drawing2D.SmoothingMode.AntiAlias;
            graphics.TextRenderingHint = SD.Text.TextRenderingHint.SingleBitPerPixel;
            graphics.DrawString(text, font, new SD.SolidBrush(SD.Color.FromArgb(255, 0, 0)), 0, 0);
            graphics.Flush();
            graphics.Dispose();
            string fileName = Path.GetFileNameWithoutExtension(Path.GetRandomFileName()) + ".jpg";
            bitmap.Save("c:\\1\\99.png", SD.Imaging.ImageFormat.Jpeg);
        }




        private void button13_Click(object sender, EventArgs e)
        {
            SD.Bitmap flag = new SD.Bitmap(1000, 500);
            SD.Graphics flagGraphics = SD.Graphics.FromImage(flag);

            SD.FontFamily fontFamily = new SD.FontFamily("GOST type A");
            SD.Font font = new SD.Font(
               fontFamily,
               32,
               SD.FontStyle.Regular,
               SD.GraphicsUnit.Pixel);
            SD.SolidBrush solidBrush = new SD.SolidBrush(SD.Color.Black);
            string string2 = @"1 Технические требования см. лист 1.1
                    2 Маркировку наносить как на стенке, так и на полках
                    клеймением и маркиратором на расстоянии 500 мм от края
                    Пример маркировки: 109-91SR301-1-B130
                    ";

            flagGraphics.TextRenderingHint = SD.Text.TextRenderingHint.AntiAlias;
            flagGraphics.DrawString(string2, font, solidBrush, new SD.PointF(10, 60));

            flag.Save("c:\\1\\333.png");
        }

        private void button14_Click(object sender, EventArgs e)
        {
            TSD.Drawing drawing = new TSD.DrawingHandler().GetActiveDrawing() as TSD.Drawing;

            TSD.DrawingObjectEnumerator allPlugins = drawing.GetSheet().GetAllObjects(typeof(TSD.Plugin));
            while (allPlugins.MoveNext())
            {
                TSD.Plugin currentPlugin = (TSD.Plugin)allPlugins.Current;
                if (currentPlugin.Name == "metcon_Revision") currentPlugin.Delete();
            }
            drawing.CommitChanges();



            //  text1_ => DESCRIPTION
            //  text2_ => INFO1
            //  text3_ => INFO2
            //  text4_ => CREATED_BY
            //  text5_ => CHECKED_BY
            //  text6_ => APPROVED_BY
            //  text7_ => DELIVERY
            //  dat1_ => DATE_CREATE
            //  dat2_ => DATE_CHECKED
            //  dat3_ => DATE_APPROVED







            TSD.DrawingObjectEnumerator sectionsDetails = drawing.GetSheet().GetAllObjects();
            foreach (TSD.DrawingObject drawingObject in sectionsDetails)
            {
                if (drawingObject is TSD.Plugin)
                {
                    var stop = "";
                    drawingObject.Delete();
                }
                if (drawingObject is TSD.Polygon)
                {
                    drawingObject.Delete();
                    drawing.CommitChanges();
                }
            }
        }

        private void Button8_Click_1(object sender, EventArgs e)
        {
            TSD.DrawingEnumerator selectedDrawings = new TSD.DrawingHandler().GetDrawingSelector().GetSelected();

            int count = 1;
            while (selectedDrawings.MoveNext())
            {
                button8.Text = count.ToString() + " из " + selectedDrawings.GetSize();
                TSD.Drawing drawing = selectedDrawings.Current as TSD.Drawing;

                TSD.DrawingObjectEnumerator objEnum = drawing.GetSheet().GetAllObjects();
                foreach (TSD.DrawingObject drawingObject in objEnum)
                {
                    if (drawingObject is TSD.Plugin)
                    {
                        TSD.Plugin plugin = drawingObject as TSD.Plugin;
                        if (plugin.Name == "COGDimensioning")
                            plugin.Delete();
                    }
                    if (drawingObject is TSD.Polygon)
                    {
                        TSD.Polygon polygon = drawingObject as TSD.Polygon;
                        string hatch = polygon.Attributes.Hatch.Name;
                        if (hatch == "ANSI31") // DIMETCOTE
                            ChangeHatchAttributes(polygon, polygon.Attributes.Hatch, "Сталь с ромбическим рифлением", 1, .5);
                        if (hatch == "Железобетон") // не грунтовать
                            ChangeHatchAttributes(polygon, polygon.Attributes.Hatch, "ANSI31", 2, 1);
                    }
                    if (drawingObject is TSD.Rectangle)
                    {
                        TSD.Rectangle rectangle = drawingObject as TSD.Rectangle;
                        string hatch = rectangle.Attributes.Hatch.Name;
                        if (hatch == "ANSI31") // DIMETCOTE
                            ChangeHatchAttributes(rectangle, rectangle.Attributes.Hatch, "Сталь с ромбическим рифлением", 1, .5);
                        if (hatch == "Железобетон") // не грунтовать
                            ChangeHatchAttributes(rectangle, rectangle.Attributes.Hatch, "ANSI31", 2, 1);
                    }
                }
                this.Update();

                drawing.CommitChanges();
                count++;
                this.Update();
            }
            ZI_library.Messages.TextMessageBox("Done!");
        }
        private void ChangeHatchAttributes(TSD.DrawingObject drawingObject,
            TSD.GraphicObjectHatchAttributes hatchAttributes, string name, double ScaleX, double ScaleY)
        {
            hatchAttributes.Name = name;
            hatchAttributes.ScaleX = ScaleX;
            hatchAttributes.ScaleY = ScaleY;
            drawingObject.Modify();
        }






        private void Button15_Click(object sender, EventArgs e)
        {
            string excelFilePath = @"\\172.16.0.16\models\2140\Отчеты\Ведомость рабочих чертежей из 1С.xlsx";

            //create a list to hold all the values
            DataTable fromExcelFileDT = new DataTable("fromExcelFile");
            fromExcelFileDT.Columns.Add("Обозначение документа", typeof(string));
            fromExcelFileDT.Columns.Add("Наименование", typeof(string));
            fromExcelFileDT.Columns.Add("Масса, кг", typeof(string));
            fromExcelFileDT.Columns.Add("Примеч.", typeof(string));

            //read the Excel file as byte array
            byte[] bin = File.ReadAllBytes(excelFilePath);

            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                ExcelWorksheet sheet = excelPackage.Workbook.Worksheets.First();

                //loop all rows
                for (int i = sheet.Dimension.Start.Row; i <= sheet.Dimension.End.Row; i++)
                {
                    DataRow row = fromExcelFileDT.NewRow();
                    if (sheet.Cells[i, 1].Value != null)
                        if (sheet.Cells[i, 1].Value.ToString() != "Итого")
                            row["Обозначение документа"] = "Заказ-КМД-" + sheet.Cells[i, 1].Value.ToString();
                        else
                            row["Обозначение документа"] = "Общая масса конструкций по чертежам, кг";
                    if (sheet.Cells[i, 2].Value != null)
                        row["Наименование"] = sheet.Cells[i, 2].Value.ToString();
                    if (sheet.Cells[i, 3].Value != null)
                        row["Масса, кг"] = sheet.Cells[i, 3].Value.ToString();
                    fromExcelFileDT.Rows.Add(row);
                }
            }
            MessageBox.Show("Done!");
        }





        private void Button16_Click(object sender, EventArgs e)
        {
            string stream = File.ReadAllText(@"C:/Users/ksd/AppData/Local/Temp/chzmk/report/DrawingListReport.html");

            var doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(stream);
            HtmlNode firstTable = doc.DocumentNode.SelectSingleNode("//table");
            var orderedCellTexts = firstTable.Descendants("tr")
                .Select(row => row.SelectNodes("th|td").Take(2).ToArray())
                .Where(cellArr => cellArr.Length == 2)
                .Select(cellArr => new { Cell1 = cellArr[0].InnerText, Cell2 = cellArr[1].InnerText })
                .OrderBy(x => x.Cell1)
                .ToList();


            var comment = doc.DocumentNode.SelectNodes("//comment()");


            foreach (HtmlAgilityPack.HtmlNode span in doc.DocumentNode.SelectNodes("//style"))
            {
                HtmlNodeCollection styles = span.ChildNodes;
                foreach (HtmlNode node in styles)
                {
                    var value = node.GetAttributeValue("border", "");
                }

                foreach (HtmlNode node in styles)
                {
                    var text = node.Attributes["border"];
                }
            }



            var body = doc.DocumentNode.SelectNodes("//body").Single();

            foreach (HtmlAgilityPack.HtmlNode td in body.SelectNodes("//td"))
            {
                var classValue = td.Attributes["class"] == null ? null : td.Attributes["class"].Value;

                var value = td.GetAttributeValue("background-color", "");

                if (classValue == "first")
                {
                    //write innerText into a table at place [i][column1]
                }
                else if (classValue == "second")
                {
                    //write innerText into the same table in [i][column2]
                }
            }


        }

        private void button17_Click(object sender, EventArgs e)
        {
            TSD.Drawing drawing = new TSD.DrawingHandler().GetActiveDrawing() as TSD.Drawing;

            int cnt = 1;

            TSD.DrawingObjectEnumerator objEnum = drawing.GetSheet().GetAllObjects();
            foreach (TSD.DrawingObject drawingObject in objEnum)
            {
                if (drawingObject is TSD.Text)
                {
                    TSD.Text text = drawingObject as TSD.Text;
                    if (text.TextString.Contains("4014"))
                    {
                        text.TextString = text.TextString.Replace("/", "-tZn/");
                        text.TextString = text.TextString + "-tZn";
                        text.Modify();
                    }
                    if (text.TextString.Contains("32484"))
                    {
                        text.TextString = text.TextString.Replace("/", "-TDG5/");
                        text.TextString = text.TextString + "-ТД5";
                        text.Modify();
                    }
                }
                Tekla.Structures.Model.Operations.Operation.DisplayPrompt(cnt + " из " + objEnum.GetSize());
                cnt++;
            }

            drawing.CommitChanges();

            ZI_library.Messages.TextMessageBox("Done!");
        }

        private void button18_Click(object sender, EventArgs e)
        {
            TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
            TSM.ModelObjectEnumerator SelectBlock = selector.GetSelectedObjects();
            SelectBlock.SelectInstances = false;
            foreach (TSM.ModelObject MO in SelectBlock)
            {
                if (MO is TSM.Grid)
                {
                    string result = "";
                    TSM.Grid grid = MO as TSM.Grid;
                    grid.Select();
                    string ZCoord = grid.CoordinateZ;
                    string[] coordinates = ZCoord.Split(' ');
                    for (int i = 0; i < coordinates.Length; i++)
                    {
                        double height = ZI_library.Math.ParseToDouble(coordinates[i]);
                        height = height / 1000;
                        string mark = height.ToString("#.000");
                        if (height == 0) mark = "0,000";
                        result += mark + " ";
                    }
                    result = result.Trim();
                    grid.LabelZ = result;
                    grid.Modify();
                    TSM.Model model = new TSM.Model();
                    model.CommitChanges();
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            model = new TSM.Model();
            TSM.Beam beam = new TSM.Beam
            {
                StartPoint = new Point(0, 0, 0),
                EndPoint = new Point(3650, 0, 0)
            };
            beam.Profile.ProfileString = "WI300-15-20*300";
            beam.Material.MaterialString = "AISI316";
            bool insert = beam.Insert();
            if (insert)
            {
                model.CommitChanges();
                Console.WriteLine("Beam inserted. " + "GUID: " + beam.Identifier.GUID.ToString());
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            Functions functions = new Functions();
            try
            {
                Dictionary<string, string> drawSOMText = new Dictionary<string, string>();

                TSM.Model model = new TSM.Model();
                string modelPath = model.GetInfo().ModelPath;
                DirectoryInfo SOMPath = new DirectoryInfo(modelPath + @"\СОМ");

                if (SOMPath.Exists)
                {
                    TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
                    TSM.ModelObjectEnumerator SelectBlock = selector.GetSelectedObjects();
                    SelectBlock.SelectInstances = false;

                    foreach (TSM.Assembly assembly in SelectBlock)
                    {
                        assembly.Select();
                        string СОМ_НазваниеФайла = "";
                        int drawID = -1;
                        assembly.GetUserProperty("metcon_SOM", ref СОМ_НазваниеФайла);

                        assembly.GetReportProperty("DRAWING.ID", ref drawID);
                        TSM.Beam Memb = new TSM.Beam
                        {
                            Identifier = new Identifier(drawID)
                        };
                        Memb.Select();
                        string drawNum = "";
                        Memb.GetUserProperty("metcon_KMD_Number", ref drawNum);
                        drawNum = "666_КМД-" + drawNum;

                        FileInfo СОМ_ПутьКФайлу = new FileInfo(SOMPath + @"\" + СОМ_НазваниеФайла + ".txt");
                        if (СОМ_ПутьКФайлу.Exists)
                        {
                            // Заполняем словарь чертеж-требование к чертежу
                            if (!drawSOMText.ContainsKey(drawNum))
                                drawSOMText.Add(drawNum, File.ReadAllText(СОМ_ПутьКФайлу.FullName));
                        }
                        else
                            MessageBox.Show("Не найден файл примечаний для " + drawNum);
                    }

                    // Сортируем по номеру чертежа
                    var sortedDic = from draw in drawSOMText
                                    let assNum = draw.Key.Split(new[] { "-" }, StringSplitOptions.RemoveEmptyEntries)
                                    orderby assNum[0].ToString(), int.Parse(assNum[1])
                                    select draw;

                    // Создаём файл excel
                    string excelFile = DateTime.Now.ToShortDateString() + ".xlsx";
                    FileInfo excelFI = new FileInfo(SOMPath + @"\" + excelFile);
                    if (excelFI.Exists)
                        excelFI.Delete();
                    using (ExcelPackage package = new ExcelPackage(excelFI))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Примечания для СОМ");

                        int startRow = 1;
                        int startCol = 1;
                        int currentRow = startRow;
                        int currentCol = startCol;

                        foreach (KeyValuePair<string, string> pair in drawSOMText)
                        {
                            currentCol = startCol;
                            worksheet.Cells[currentRow, currentCol].Value = pair.Key;
                            currentCol++;
                            worksheet.Cells[currentRow, currentCol].Value = pair.Value;
                            currentRow++;
                        }
                        worksheet.Cells.AutoFitColumns();
                        worksheet.Calculate();
                        package.Save();
                    }
                    ZI_library.Reports.OpenFile(excelFI.FullName);
                }
                else
                    MessageBox.Show("Создайте папку СОМ и поместите в неё файлы, содержащие текст примечаний для СОМ!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            TSM.Model model = new TSM.Model();
            TSM.ModelInfo minfo = model.GetInfo();
            TeklaStructuresFiles tsf = new TeklaStructuresFiles(minfo.ModelPath);
            List<string> PropertyFileDirectories = tsf.PropertyFileDirectories;

            string templatesPath = string.Empty;
            TeklaStructuresSettings.GetAdvancedOption("XS_TEMPLATE_DIRECTORY", ref templatesPath);
            string templates1Path = string.Empty;
            TeklaStructuresSettings.GetAdvancedOption("XS_TEMPLATE_DIRECTORY_SYSTEM", ref templates1Path);
            string modelPath = model.GetInfo().ModelPath;


            string report1 = Tekla.Structures.Dialog.UIControls.EnvironmentVariables.GetEnvironmentVariable("XS_TEMPLATE_DIRECTORY");
            string report = Tekla.Structures.Dialog.UIControls.EnvironmentVariables.GetEnvironmentVariable("XS_TEMPLATE_DIRECTORY_SYSTEM");



            FileInfo atribFile = Tekla.Structures.Dialog.UIControls.EnvironmentFiles.GetAttributeFile("boltsdata.exe.config");

            atribFile = Tekla.Structures.Dialog.UIControls.EnvironmentFiles.GetAttributeFile("ALNG2ProjectsData.metcon");

            List<string> stdProp = Tekla.Structures.Dialog.UIControls.EnvironmentFiles.GetMultiDirectoryFileList("rpt");

            List<string> stdProp1 = Tekla.Structures.Dialog.UIControls.EnvironmentFiles.GetMultiDirectoryFileList("xml.rpt");
            stdProp1.Sort();

            string path = Path.Combine(@"\\newbuffer\buffer\SAPR\TechnicalRequests\", "000000113", "bb569eed-4b65-4729-b40c-5bb543d50280", "Questions");


            List<string> list = tsf.GetMultiDirectoryFileList("xml", false);

            var fileee = Tekla.Structures.Dialog.UIControls.EnvironmentFiles.GetAttributeFile("objects.inp");
        }








        private void button22_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    TSD.DrawingEnumerator drawingEnumerator = new TSD.DrawingHandler().GetDrawingSelector().GetSelected();

            //    while (drawingEnumerator.MoveNext())
            //    {
            //        var selDraw = drawingEnumerator.Current;
            //        string assNumber = selDraw.Mark.Replace("[","").Replace("]","").Replace(".", "-");


            //        PropertyInfo pi = drawingEnumerator.Current.GetType()
            //            .GetProperty("Identifier", BindingFlags.Instance | BindingFlags.NonPublic);
            //        object value = pi.GetValue(drawingEnumerator.Current, null);
            //        Identifier Identifier = (Identifier)value;
            //        TSM.Beam fakebeam = new TSM.Beam();
            //        fakebeam.Identifier = Identifier;
            //        string drawType = "";
            //        fakebeam.GetReportProperty("TYPE", ref drawType);

            //        string isp = "";
            //        selDraw.GetUserProperty("metcon_KMD_Ispolnil", ref isp);

            //        string baseIsp = "";

            //        switch (isp)
            //        {
            //            case "Корнилова":
            //                baseIsp = "kornilova";
            //                break;
            //            case "Зуйков":
            //                baseIsp = "zuiay";
            //                break;
            //            case "Матвеева":
            //                baseIsp = "matov";
            //                break;
            //            case "Самсонов":
            //                baseIsp = "samaa";
            //                break;
            //            default:
            //                break;
            //        }

            //        string responseString = DBQueryFree("UPDATE public.model_info_planning SET drawing_developer='"
            //            + baseIsp + "' WHERE assembly_number='" + assNumber + "'");
            //    }
            //}

            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
            try
            {

                string responseString = DBQueryFree(TB_Query.Text);
                JObject jObject = JObject.Parse(responseString);
                var msg = jObject.SelectToken("msg");

                string login = string.Empty;
                if (msg.HasValues)
                    login = (string)msg.Parent.First[0]["login"];
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private static string DBQueryFree(string query)
        {
            string adress = "http://tsms:49074/qIGE41lSpFFC2kY9eKhg2HycG?query=" + ZI_library.StringExtensions.EncodeBase64(query);
            return DBQuery(adress);
        }
        private static string DBQuery(string adress)
        {
            string responseString = String.Empty;
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(adress);
                var response = (HttpWebResponse)request.GetResponse();

                using (var stream = response.GetResponseStream())
                {
                    using (var sreader = new StreamReader(stream))
                    {
                        responseString = sreader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return responseString;
        }







        private void button23_Click(object sender, EventArgs e)
        {
            string responseString = DBQueryFree("select *, to_char(create_date, 'dd.MM.yyyy HH24:MI') as formatteddate from public.model_info where ordergroupcode='000000113' ORDER BY numinproject");

            JObject obj = JObject.Parse(responseString);
            string msg = obj.SelectToken("msg").ToString();
            DataTable questionsDT = new DataTable("Questions");

            questionsDT = (DataTable)JsonConvert.DeserializeObject(msg, (typeof(DataTable)));

        }


        private void button35_Click(object sender, EventArgs e)
        {
            string projResponse = DBQueryFree("select id from public.pkoservices_projects_book where title='" + "2680" + "'");
            JObject jObject = JObject.Parse(projResponse);
            var msg = jObject.SelectToken("msg");

            int projectID = int.MinValue;
            if (msg.HasValues)
                projectID = (int)msg.Parent.First[0]["id"];
        }



        public class QuestionInfo
        {
            public string Id { get; set; }
            public string project_name { get; set; }
            public int NumInProject { get; set; }
            public int NumInOrderGroup { get; set; }
            public string TitleRus { get; set; }
            public string TitleEng { get; set; }
            public string ContentRus { get; set; }
            public string ContentEng { get; set; }
            public string create_date { get; set; }
            public string FormattedDate { get; set; }
            public string Username { get; set; }
            public string EnvironmentUsername { get; set; }
            public int Status { get; set; }
            public string StatusString { get; set; }
            public string Submodel { get; set; }
            public string Comment { get; set; }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            string[] values = Enumerable.Range(0, 27 * 26).Select(n => new String(new[]
            {
                (char)('@' + n / 26), (char)('A' + n % 26)
            },
                n < 26 ? 1 : 0, n < 26 ? 1 : 2)).ToArray();

            string[] result = SubArray(values, 26, values.Length - 26);

            //int index = value + 26;
            //string result = string.Empty;
            //while (--index >= 0)
            //{
            //    result = (char)('A' + index % 26) + result;
            //    index /= 26;
            //}
            //return result;
        }
        public T[] SubArray<T>(T[] data, int index, int length)
        {
            T[] result = new T[length];
            Array.Copy(data, index, result, 0, length);
            return result;
        }




        private void button25_Click(object sender, EventArgs e)
        {
            TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
            TSM.ModelObjectEnumerator SelectBlock = selector.GetSelectedObjects();
            SelectBlock.SelectInstances = false;

            TSM.Part selectedPart = null;
            if (SelectBlock.GetSize() == 1)
                while (SelectBlock.MoveNext())
                    selectedPart = SelectBlock.Current as TSM.Part;

            TSM.Assembly assembly = selectedPart.GetAssembly();
            ArrayList secondaries = assembly.GetSecondaries();
            secondaries.Add(assembly.GetMainPart() as TSM.Part);

            TSM.Solid solidSelectedPart = selectedPart.GetSolid(TSM.Solid.SolidCreationTypeEnum.HIGH_ACCURACY);

            int counter = 0;
            for (int sec = 0; sec < secondaries.Count; sec++)
            {
                TSM.Part curPart = secondaries[sec] as TSM.Part;

                // убираем выбранную деталь из перебора
                selectedPart.Select();
                if (selectedPart.Identifier.GUID != curPart.Identifier.GUID)
                {
                    List<ContactPlane> contactPlanes = ContactParts.GetContactPlanesOfParts(selectedPart, curPart, 2);
                    if (contactPlanes.Count > 0)
                    {
                        counter += contactPlanes.Count;

                        for (int cp = 0; cp < contactPlanes.Count; cp++)
                        {
                            ContactPlane contactPlane = contactPlanes[cp];

                            for (int i = 0; i < contactPlane.FirstFaceVertex.Count; i++)
                            {
                                Point point = contactPlane.FirstFaceVertex[i];
                                GraphDrawer.graphicsDrawer.DrawText(point, i.ToString(), GraphDrawer.blue);



                                Point proj = Projection.PointToPlane(point,
                                    new GeometricPlane(contactPlane.SecondFaceVertex[0], contactPlane.SecondFace.Normal));
                                GraphDrawer.graphicsDrawer.DrawText(proj, "proj" + i.ToString(), GraphDrawer.green);

                                TSM.Polygon polygon = new TSM.Polygon();
                                polygon.Points = new ArrayList(contactPlane.SecondFaceVertex);
                                for (int polygonPoint = 0; polygonPoint < polygon.Points.Count; polygonPoint++)
                                {
                                    Point polPoint = polygon.Points[polygonPoint] as Point;
                                    GraphDrawer.graphicsDrawer.DrawText(polPoint, "polygon" + polygonPoint.ToString(), GraphDrawer.yellow);
                                }
                                bool pointInside = cs_net_lib.Geo.IsPointInsidePolygon2D(polygon, point);

                                Point closestPoint = cs_net_lib.Geo.GetClosestPointBetweenPointAndLineSegment3D(point, polygon.Points[0] as Point, polygon.Points[1] as Point);
                                GraphDrawer.graphicsDrawer.DrawText(closestPoint, "closestPoint", GraphDrawer.purple);


                            }
                            for (int i = 0; i < contactPlane.SecondFaceVertex.Count; i++)
                            {
                                Point point = contactPlane.SecondFaceVertex[i];
                                GraphDrawer.graphicsDrawer.DrawText(point, i.ToString(), GraphDrawer.red);
                            }






                        }
                    }
                }
            }
            Messages.TextMessageBox("Граней: " + counter.ToString());
        }



        internal static bool ArePointAligned(Point point1, Point point2, Point point3)
        {
            Vector vector1 = new Vector(point2.X - point1.X, point2.Y - point1.Y, point2.Z - point1.Z);
            Vector vector2 = new Vector(point3.X - point1.X, point3.Y - point1.Y, point3.Z - point1.Z);
            return Tekla.Structures.Geometry3d.Parallel.VectorToVector(vector1, vector2);
        }

        private void button26_Click(object sender, EventArgs e)
        {
            string rawFilter = "dummy,      CG, шлак";
            string[] filterArr = rawFilter.Split(new char[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            string filter = "";

            for (int i = 0; i < filterArr.Length; i++)
            {
                filterArr[i] = filterArr[i].Trim();
                filter += "НомерПозиции" + " NOT LIKE '" + filterArr[i] + "*' AND ";
            }
            filter = filter.TrimEnd();
            filter = filter.TrimEnd(new char[] { ' ', 'A', 'N', 'D' });
            MessageBox.Show(filter);
        }






        private void btn_sym_Click(object sender, EventArgs e)
        {
            //uint unicodeValue = (uint)tb_sym.Text[0];
            //lbl_sym.Text = unicodeValue.ToString("X");

            //var enumerator = StringInfo.GetTextElementEnumerator(tb_sym.Text);
            //while (enumerator.MoveNext())
            //{
            //    Console.WriteLine(enumerator.Current);
            //}

            //string path = @"C:\Windows\System32\getuname.dll";
            //using (var reader = new Win32ResourceReader(path))
            //{
            //    string name = reader.GetString(unicodeValue);
            //    MessageBox.Show(name);
            //}

            ////MessageBox.Show(CharUnicodeInfo.GetUnicodeCategory(tb_sym.Text,0).ToString());
            MessageBox.Show(is_only_eng_letters_and_digits(tb_sym.Text).ToString());
        }
        private static bool is_only_eng_letters_and_digits(string str)
        {
            foreach (char ch in str)
            {
                if (!(ch >= 'A' && ch <= 'Z') && !(ch >= 'a' && ch <= 'z') && !(ch >= '0' && ch <= '9'))
                {
                    return false;
                }
            }
            return true;
        }



        public class Win32ResourceReader : IDisposable
        {
            private IntPtr _hModule;

            public Win32ResourceReader(string filename)
            {
                _hModule = LoadLibraryEx(filename, IntPtr.Zero, LoadLibraryFlags.AsDataFile | LoadLibraryFlags.AsImageResource);
                if (_hModule == IntPtr.Zero)
                    throw Marshal.GetExceptionForHR(Marshal.GetHRForLastWin32Error());
            }

            public string GetString(uint id)
            {
                var buffer = new StringBuilder(1024);
                LoadString(_hModule, id, buffer, buffer.Capacity);
                if (Marshal.GetLastWin32Error() != 0)
                    throw Marshal.GetExceptionForHR(Marshal.GetHRForLastWin32Error());
                return buffer.ToString();
            }

            ~Win32ResourceReader()
            {
                Dispose(false);
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            public void Dispose(bool disposing)
            {
                if (_hModule != IntPtr.Zero)
                    FreeLibrary(_hModule);
                _hModule = IntPtr.Zero;
            }

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hReservedNull, LoadLibraryFlags dwFlags);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            static extern bool FreeLibrary(IntPtr hModule);

            [Flags]
            enum LoadLibraryFlags : uint
            {
                AsDataFile = 0x00000002,
                AsImageResource = 0x00000020
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            TSM.ModelObjectEnumerator enumerator = new TSM.Model().GetModelObjectSelector().GetAllObjects();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            try
            {
                TSM.Model model = new TSM.Model();

                if (model.GetConnectionStatus())
                {
                    TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
                    TSM.ModelObjectEnumerator SelectBlock = selector.GetSelectedObjects();
                    SelectBlock.SelectInstances = false;

                    if (SelectBlock.GetSize() > 0)
                    {
                        while (SelectBlock.MoveNext())
                        {
                            TSM.ModelObject modelObject = SelectBlock.Current as TSM.ModelObject;
                            MessageBox.Show(modelObject.GetType().ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

















        private void button29_Click(object sender, EventArgs e)
        {
            //Functions functions = new Functions();
            //TSM.Model model = new TSM.Model();

            //// Создаём отчёт
            //DataSet ReportDataSet = ZI_library.Reports.ReportDataSet(model, "0_ALNG_EXCEL_Import.xml", false, "", "", "");

            //if (ReportDataSet == null)
            //{
            //    MessageBox.Show("Отчёт не создан!");
            //    return;
            //}
            //else
            //{
            //    try
            //    {
            //        metcon_ALNG2_reports_onshore.Reports repOnshore = new metcon_ALNG2_reports_onshore.Reports();
            //        metcon_ALNG2_reports_onshore.Reports.ErrorMessage error = new metcon_ALNG2_reports_onshore.Reports.ErrorMessage();
            //        repOnshore.MakeExcelImportFiles(ReportDataSet, out error, out FileInfo assFilename, out FileInfo partfilename);
            //        if (!error.Status)
            //        {
            //            MessageBox.Show(error.Message);
            //        }
            //        else
            //        {
            //            // Открываем созданные файлы
            //            functions.OpenFile(assFilename.FullName);
            //            functions.OpenFile(partfilename.FullName);
            //        }
            //    }
            //    catch
            //    {

            //    }
            //}
        }






















        private void button30_Click(object sender, EventArgs e)
        {
            TSD.DrawingHandler DrawingHandler = new TSD.DrawingHandler();

            if (DrawingHandler.GetConnectionStatus())
            {
                TSD.Drawing CurrentDrawing = DrawingHandler.GetActiveDrawing();
                if (CurrentDrawing != null)
                {
                    TSD.DrawingObjectEnumerator markList = CurrentDrawing.GetSheet().GetAllObjects(typeof(TSD.Text));
                    while (markList.MoveNext())
                    {
                        TSD.Text currentMark = markList.Current as TSD.Text;

                        if (currentMark != null & currentMark.TextString.Contains(dwgfrom.Text))
                        {
                            string newText = currentMark.TextString.Replace(dwgfrom.Text, dwgto.Text);
                            currentMark.Select();
                            currentMark.TextString = newText;
                            currentMark.Modify();
                        }
                    }
                    CurrentDrawing.CommitChanges();
                }
            }
        }




        private void button31_Click(object sender, EventArgs e)
        {


            string tempFolder = Path.GetTempPath();
            tempFolder = @"\\fileserver\buffer\SAPR\TechnicalRequests\000000113\f51fd72a-273a-4f17-a73a-da92ca1fda93\Questions";
            string answerpath = Path.Combine(tempFolder, "Тестовый запрос.docx");




            FileInfo sourceFile = new FileInfo(answerpath);
            string destFile = answerpath.Replace(sourceFile.Extension, ".pdf");
            FileInfo destFI = new FileInfo(destFile);

            bool CreateDBEntry = true;
            if (destFI.Exists)
            {
                //destFI.Delete();
                CreateDBEntry = false;
            }

            System.Threading.Thread.Sleep(1000);

            //Load Document  
            Spire.Doc.Document document = new Spire.Doc.Document();
            document.LoadFromFile(sourceFile.FullName);
            //Convert Word to PDF  
            document.SaveToFile(destFile, Spire.Doc.FileFormat.PDF);
            document.Close();
            document.Dispose();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    while (drawingEnumerator.MoveNext())
            //    {
            //        var selDraw = drawingEnumerator.Current;
            //        string name_base = selDraw.Mark;

            //        string drawType = "";
            //        if (selDraw is TSD.AssemblyDrawing)
            //            drawType = "A";
            //        if (selDraw is TSD.SinglePartDrawing)
            //            drawType = "W";
            //        if (selDraw is TSD.MultiDrawing)
            //            drawType = "M";
            //        if (selDraw is TSD.GADrawing)
            //            drawType = "G";
            //        if (selDraw is TSD.CastUnitDrawing)
            //            drawType = "C";

            //        //PropertyInfo pi = drawingEnumerator.Current.GetType().GetProperty("Identifier", BindingFlags.Instance | BindingFlags.NonPublic);
            //        //object value = pi.GetValue(drawingEnumerator.Current, null);
            //        //Identifier Identifier = (Identifier)value;
            //        //TSM.Beam fakebeam = new TSM.Beam();
            //        //fakebeam.Identifier = Identifier;
            //        //string drawType = "";
            //        //fakebeam.GetReportProperty("TYPE", ref drawType);

            //        string num = "";
            //        selDraw.GetUserProperty("metcon_KMD_Number", ref num);
            //        string list = "";
            //        selDraw.GetUserProperty("metcon_KMD_List_№", ref list);
            //        string listov = "";
            //        selDraw.GetUserProperty("metcon_KMD_Listov", ref listov);

            //        DrawCompare drawing = new DrawCompare()
            //        {
            //            draw_type = drawType,
            //            name_base = name_base,
            //            kmd_number_tekla = num,
            //            list_tekla = list,
            //            listov_tekla = listov,
            //            tekla_drawing = selDraw
            //        };
            //        DrawingsTeklaPlusDB.Add(drawing);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}
        }

        private void button33_Click(object sender, EventArgs e)
        {
            try
            {
                //    //metcon_NetStandardLibrary.DrawMiss drawMiss = new metcon_NetStandardLibrary.DrawMiss(new metcon_NetStandardLibrary.ApplicationContext());
                //    //metcon_NetStandardLibrary.Errors errors = new metcon_NetStandardLibrary.Errors();
                //    //List<metcon_NetStandardLibrary.DrawMiss> list = drawMiss.MissingNumbersList("8b250cd2-a518-4e78-8c3d-12c0a73e5e95", out errors);

                //metcon_NetStandardLibrary.DrawingEnum.DrawingDB savedDraw = drawingEnum.AddDrawNumbersToDBFunction("71b5cb4d-9357-4044-98aa-5d857d7f8eae", out adderrors);

                //metcon_NetStandardLibrary.ApplicationContext db=  new metcon_NetStandardLibrary.ApplicationContext();

                //List<metcon_NetStandardLibrary.DrawingEnum.DrawingDB> DrawingsList =
                //    db.model_drawing_enum.FromSql(@"SELECT model_drawing_enum.*, pkoservices_projects_book.title
                //    FROM model_drawing_enum
                //    JOIN pkoservices_projects_book ON model_drawing_enum.project_id = pkoservices_projects_book.id
                //    WHERE model_drawing_enum.model_id = {0}", ProjectGUID).ToList();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            TSM.Model model = new TSM.Model();

            if (model.GetConnectionStatus())
            {
                TSM.UI.ModelObjectSelector selector = new TSM.UI.ModelObjectSelector();
                TSM.ModelObjectEnumerator SelectBlock = selector.GetSelectedObjects();
                SelectBlock.SelectInstances = false;

                if (SelectBlock.GetSize() > 0)
                {
                    while (SelectBlock.MoveNext())
                    {
                        TSM.ModelObject modelObject = SelectBlock.Current;
                        modelObject.Select();
                        modelObject.SetUserProperty("metcon_Ochered", "123");
                        modelObject.SetUserProperty("ES_DESC", -1);
                    }
                    MessageBox.Show("Готово!");
                }
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            var newGuid = Guid.NewGuid().ToString();
            string base64Guid = Convert.ToBase64String(Guid.NewGuid().ToByteArray());
        }











        public DataTable AttachmentsDT;
        // Получение почты
        private void button37_Click(object sender, EventArgs e)
        {
            try
            {
                using (var client = new ImapClient())
                {
                    client.ServerCertificateValidationCallback = (mysender, certificate, chain, sslPolicyErrors) => { return true; };
                    client.Connect("ms.metcon.ru", 143, SecureSocketOptions.StartTls);
                    client.Authenticate("robot-pko@metcon.ru", "robot-pkometcon");
                    client.Inbox.Open(FolderAccess.ReadOnly);

                    string projectID = "ProjectID=17";
                    string requestID = "RequestID=2902";

                    // search for messages where ... equals ... and ...
                    var query = SearchQuery.BodyContains(projectID).And(SearchQuery.BodyContains(requestID));
                    var uids = client.Inbox.Search(query);

                    //((UniqueIdSet)uids).Reverse();


                    // fetch summary information for the search results (we will want the UID and the BODYSTRUCTURE
                    // of each message so that we can extract the text body and the attachments)
                    var items = client.Inbox.Fetch(uids, MessageSummaryItems.UniqueId
                        | MessageSummaryItems.BodyStructure
                        | MessageSummaryItems.InternalDate);

                    var newitems = items.OrderByDescending(x => x.InternalDate).ToList();

                    foreach (var item in newitems)
                    {
                        // determine a directory to save stuff in
                        var directory = Path.Combine(@"D:\TEMP\Mailkit", item.UniqueId.ToString());
                        // create the directory
                        Directory.CreateDirectory(directory);
                        // IMessageSummary.TextBody is a convenience property that finds the 'text/plain' body part for us
                        var bodyPart = item.TextBody;
                        // download the 'text/plain' body part
                        var body = (TextPart)client.Inbox.GetBodyPart(item.UniqueId, bodyPart);
                        // TextPart.Text is a convenience property that decodes the content and converts the result to a string for us
                        var text = body.Text;
                        File.WriteAllText(Path.Combine(directory, "body.txt"), text);

                        //var htmlBodyPart = item.HtmlBody;
                        //var mimeEntity = client.Inbox.GetBodyPart(item.Index, item.HtmlBody);

                        var textBodyPart = item.TextBody;
                        var mimeEntity = client.Inbox.GetBodyPart(item.Index, item.TextBody);
                        string test = ((TextPart)mimeEntity).Text;
                        string test1 = ((MessageSummary)item).InternalDate.Value.LocalDateTime.ToString();

                        AttachmentsDT = new DataTable("Attachments");

                        DataColumn saveToDB = new DataColumn("Save", typeof(bool));
                        DataColumn fileName = new DataColumn("Filename", typeof(string));
                        DataColumn sourceFilePath = new DataColumn("SourceFilePath", typeof(string));
                        DataColumn destFilePath = new DataColumn("DestFilePath", typeof(string));
                        AttachmentsDT.Columns.Add(saveToDB);
                        AttachmentsDT.Columns.Add(fileName);
                        AttachmentsDT.Columns.Add(sourceFilePath);
                        AttachmentsDT.Columns.Add(destFilePath);

                        // now iterate over all of the attachments and save them to disk
                        foreach (var attachment in item.Attachments)
                        {
                            // download the attachment just like we did with the body
                            var entity = client.Inbox.GetBodyPart(item.UniqueId, attachment);

                            DataRow row = null;

                            // attachments can be either message/rfc822 parts or regular MIME parts
                            if (entity is MessagePart)
                            {
                                var rfc822 = (MessagePart)entity;
                                var path = Path.Combine(directory, attachment.FileName);
                                rfc822.Message.WriteTo(path);

                                row = AttachmentsDT.NewRow();
                                row.ItemArray = new object[] { true, attachment.FileName, path + ".eml",
                                        Path.Combine( @"\\fileserver\buffer\SAPR\TechnicalRequests\000000113\dc175c46-585d-4a2b-9eb7-774ccfb01f6d\Answers",
                                        "3"+"_"+ attachment.FileName) };
                                AttachmentsDT.Rows.Add(row);
                            }
                            else
                            {
                                //var part = (MimePart)entity;
                                //// note: it's possible for this to be null, but most will specify a filename
                                //var fileName = part.FileName;
                                //var path = Path.Combine(directory, fileName);
                                //// decode and save the content to a file
                                //using (var stream = File.Create(path))
                                //    part.Content.DecodeTo(stream);

                                var part = (MimePart)entity;
                                // note: it's possible for this to be null, but most will specify a filename
                                row = AttachmentsDT.NewRow();
                                row.ItemArray = new object[] { true, part.FileName };
                                AttachmentsDT.Rows.Add(row);
                                //var path = System.IO.Path.Combine(directory, fileName);
                                //// decode and save the content to a file
                                //using (var stream = File.Create(path))
                                //part.Content.DecodeTo(stream);
                            }
                        }
                    }

                    client.Disconnect(true);


                    //var subFolders = client.Inbox.GetSubfolders();
                    //var toplevel = client.GetFolder(client.PersonalNamespaces[0]);
                    //foreach (var folder in toplevel.GetSubfolders())
                    //{
                    //    if (folder.Name == "mailkit")
                    //    {
                    //        folder.Delete();
                    //        continue;
                    //    }
                    //}

                    //client.Inbox.Open(FolderAccess.ReadWrite);
                    //var newSubFolder = client.Inbox.Create("44", true);
                    //newSubFolder.Subscribe();

                    //foreach (var folder in client.Inbox.GetSubfolders())
                    //{
                    //    if (folder.Name == "123")
                    //        folder.Delete();
                    //}

                    //foreach (var uid in uids)
                    //{
                    //    var message = client.Inbox.GetMessage(uid);
                    //    write the message to a file
                    //    message.WriteTo(string.Format("{0}.eml", uid));
                    //}
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }









        private void button38_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                sb.AppendLine("<td style=\"display:none!important;");
                sb.AppendLine("visibility:hidden;");
                sb.AppendLine("mso-hide:all;");
                sb.AppendLine("font-size:1px;");
                sb.AppendLine("color:#ffffff;");
                sb.AppendLine("line-height:1px;");
                sb.AppendLine("max-height:0px;");
                sb.AppendLine("max-width:0px;");
                sb.AppendLine("opacity:0;");
                sb.AppendLine("overflow:hidden;\">");
                sb.AppendLine("This is preheader text.");
                sb.AppendLine("</td>");


                sb.AppendLine("<b>Номер запроса:</b> " + "Номер запроса" + "<br />");
                sb.AppendLine("<b>Заголовок запроса:</b> " + "Заголовок запроса" + "<br />");
                sb.AppendLine("<b>Текст запроса:</b> " + "Текст запроса" + "<br />");

                sb.AppendLine("<p style=\"display:none;\">" + "13579" + "</p>");

                SendEmail(new List<string>() { "sk_zlat@mail.ru" }, "Запрос по объекту: <p style=\"display:none;\">weewfewfewfewf</p>" + "Объект", sb.ToString());

                // Убедительная просьба отвечать на каждое письмо
                // Добавить ссылку для запроса реестра. Илью попросить сделать службу
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public void SendEmail(List<string> emails, string subject, string message)
        {
            var emailMessage = new MimeMessage();

            emailMessage.From.Add(new MailboxAddress("Сервисы ПКО", "robot-pko@metcon.ru"));
            if (emails != null)
                foreach (var item in emails) emailMessage.To.Add(MailboxAddress.Parse(item));
            emailMessage.Subject = subject;

            var builder = new BodyBuilder
            {
                HtmlBody = message
            };

            emailMessage.Body = builder.ToMessageBody();

            // Отправляем письмо
            var IsEmailSent = false;
            try
            {
                using (var client = new SmtpClient())
                {
                    client.ServerCertificateValidationCallback = (mysender, certificate, chain, sslPolicyErrors) => { return true; };
                    client.Connect("ms.metcon.ru", 587, SecureSocketOptions.StartTls);
                    client.Authenticate("robot-pko@metcon.ru", "robot-pkometcon");

                    client.Send(emailMessage);
                    IsEmailSent = true;
                    client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            // Сохраняем отправленное письмо в папке "Отправленные"
            if (IsEmailSent)
            {
                using (var client = new ImapClient())
                {
                    client.ServerCertificateValidationCallback = (mysender, certificate, chain, sslPolicyErrors) => { return true; };
                    client.Connect("ms.metcon.ru", 143, SecureSocketOptions.StartTls);
                    client.Authenticate("robot-pko@metcon.ru", "robot-pkometcon");

                    var inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadWrite);

                    IMailFolder SentFolder = client.GetFolder("Sent");
                    SentFolder.Open(FolderAccess.ReadWrite);
                    SentFolder.Append(emailMessage);
                }
            }
        }


        private void button39_Click(object sender, EventArgs e)
        {
            var date = TxDate.ConvertToWindowsDate(1646092800);
            var toDB = string.Join("-", date.Day, date.Month, date.Year);
            MessageBox.Show(toDB);
        }

        private void button40_Click(object sender, EventArgs e)
        {
            string[] array = { "Б-11", "Б-1", "Б-10", "Б-6", "Б4-11", "Б4-1" };
            bool flag = true;
            while (flag)
            {
                flag = false;
                for (int i = 0; i < array.Length - 1; ++i)
                    if (GetSortString(array[i]).CompareTo(GetSortString(array[i + 1])) > 0)
                    {
                        string buf = array[i];
                        array[i] = array[i + 1];
                        array[i + 1] = buf;
                        flag = true;
                    }
            }
            foreach (string s in array)
                Console.WriteLine("{0} ", s);
        }

        private string GetSortString(string s)
        {
            string sOut = s;

            string Num = "";
            string Str = "";
            int addZ = 4;

            if (sOut.IndexOf(" - ") > -1)
            {
                addZ = 4 - (sOut.Length - sOut.IndexOf(" - ") - 3);
                sOut = sOut.Replace(" - ", "");
            }

            foreach (char Nm in sOut)
            {
                if (!char.IsNumber(Nm))
                    Str += Nm;
                else
                    Num += Nm;
            }
            sOut = Str + Num.PadLeft(20 + (4 - addZ), '0') + "".PadRight(addZ, '0');
            return sOut;
        }

        private void button41_Click(object sender, EventArgs e)
        {
            // Creates a selection filter for the following filter expression:
            // (PartName == BEAM1 OR PartName == BEAM2 OR PartName == BEAM3 OR PartComment StartsWith test)

            // Creates the filter expressions
            PartFilterExpressions.Name PartName = new PartFilterExpressions.Name();
            StringConstantFilterExpression Beam1 = new StringConstantFilterExpression("ЦЙ-1");
            StringConstantFilterExpression Beam2 = new StringConstantFilterExpression("BEAM2");
            StringConstantFilterExpression Beam3 = new StringConstantFilterExpression("BEAM3");

            // Creates a custom part filter
            PartFilterExpressions.CustomString PartComment = new PartFilterExpressions.CustomString("Comment");
            StringConstantFilterExpression Test = new StringConstantFilterExpression("test");

            // Creates the binary filter expressions
            BinaryFilterExpression Expression1 = new BinaryFilterExpression(PartName, StringOperatorType.IS_EQUAL, Beam1);
            BinaryFilterExpression Expression2 = new BinaryFilterExpression(PartName, StringOperatorType.IS_EQUAL, Beam2);
            BinaryFilterExpression Expression3 = new BinaryFilterExpression(PartName, StringOperatorType.IS_EQUAL, Beam3);
            BinaryFilterExpression Expression4 = new BinaryFilterExpression(PartComment, StringOperatorType.STARTS_WITH, Test);

            // Creates the binary filter expression collection
            BinaryFilterExpressionCollection ExpressionCollection = new BinaryFilterExpressionCollection();
            ExpressionCollection.Add(new BinaryFilterExpressionItem(Expression1, BinaryFilterOperatorType.BOOLEAN_OR));
            ExpressionCollection.Add(new BinaryFilterExpressionItem(Expression2, BinaryFilterOperatorType.BOOLEAN_OR));
            ExpressionCollection.Add(new BinaryFilterExpressionItem(Expression3, BinaryFilterOperatorType.BOOLEAN_OR));
            ExpressionCollection.Add(new BinaryFilterExpressionItem(Expression4));

            string AttributesPath = Path.Combine(@"D:\TEMP\Filter");
            string FilterName = Path.Combine(AttributesPath, "filter");

            Filter Filter = new Filter(ExpressionCollection);
            // Generates the filter file
            Filter.CreateFile(FilterExpressionFileType.OBJECT_GROUP_VIEW, FilterName);  // This line is modified

        }








        public class DrawingDataFrom1C
        {
            public DrawingDataHead Head { get; set; }
            public string Manager { get; set; }
            public int CommonTypes { get; set; }
            /// <summary>
            /// Столбцы основной таблицы
            /// </summary>
            public string[] TableHead { get; set; }
            /// <summary>
            /// Основная таблица
            /// </summary>
            public string[][] Table { get; set; }
        }

        public class DrawingDataHead
        {
            public int Code { get; set; }
            public string Message { get; set; }
            public string Description { get; set; }
            public string OrderClosingDate { get; set; }
        }

        public class Drawing1C
        {
            public int IsXDraw { get; set; }
            public string Weight { get; set; }
            public string DateOfEntry { get; set; }
            public string TypeOfNomenclature { get; set; }
        }

        public class Dates
        {
            public DateTime MinDate { get; set; }
            public DateTime MaxDate { get; set; }
        }

        string responseString;
        JObject obj;
        string msg;

        Dictionary<string, double> ProjectsWeightDic;
        Dictionary<string, List<Drawing1C>> ProjectsDrawings1C;
        Dictionary<string, Dates> ProjectsDates1C;


        private void ProjectsFrom1C_Click(object sender, EventArgs e)
        {
            ProjectsWeightDic = new Dictionary<string, double>();
            ProjectsDrawings1C = new Dictionary<string, List<Drawing1C>>();
            ProjectsDates1C = new Dictionary<string, Dates>();

            char[] charSeparators = new char[] { ' ' };
            string[] projects = TB_ProjectsList.Text.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < projects.Length; i++)
            {
                List<Drawing1C> projectDrawings = GetDrawingDataFrom1C(projects[i]);
                ProjectsDrawings1C.Add(projects[i], projectDrawings);
            }

            FileInfo reportFile = new FileInfo(@"\\172.16.0.23\models\Temp\Итоги по заказам\Итоги КБ4.xlsx");
            File.Delete(reportFile.FullName);
            string SheetAssName = "КБ4";
            int currentRow;
            int currentCol;
            using (ExcelPackage package = new ExcelPackage(reportFile))
            {
                ExcelWorksheet testWorksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == SheetAssName);
                ExcelWorksheet worksheet = null;
                if (testWorksheet == null)
                    worksheet = package.Workbook.Worksheets.Add(SheetAssName);
                else
                    worksheet = package.Workbook.Worksheets[SheetAssName];

                int startRow = 1;
                int startCol = 1;

                currentRow = startRow;
                currentCol = startCol;


                worksheet.Cells[currentRow, 1].Value = "Заказ";
                worksheet.Cells[currentRow, 2].Value = "Масса, кг";
                worksheet.Cells[currentRow, 3].Value = "Начало запуска";
                worksheet.Cells[currentRow, 4].Value = "Окончание запуска";



                currentRow++;
                currentCol = startCol;

                foreach (KeyValuePair<string, double> projectPair in ProjectsWeightDic)
                {
                    currentCol = startCol;
                    worksheet.Cells[currentRow, currentCol].Value = projectPair.Key;
                    currentCol++;
                    worksheet.Cells[currentRow, currentCol].Value = projectPair.Value;

                    Dates dates = new Dates();
                    ProjectsDates1C.TryGetValue(projectPair.Key, out dates);

                    // Минимальная дата
                    currentCol++;
                    worksheet.Cells[currentRow, currentCol].Value = dates.MinDate.ToShortDateString();

                    // Максимальная дата
                    currentCol++;
                    worksheet.Cells[currentRow, currentCol].Value = dates.MaxDate.ToShortDateString();

                    currentRow++;
                }

                worksheet.Cells.AutoFitColumns();
                worksheet.Calculate();
                package.Save();
            }

            Reports.OpenFile(reportFile.FullName);
        }

        public List<Drawing1C> GetDrawingDataFrom1C(string project)
        {
            try
            {
                List<Drawing1C> drawingFrom1CList = new List<Drawing1C>();
                Functions functions = new Functions();
                DrawingDataFrom1C drawingDataFrom1C = new DrawingDataFrom1C();


                string textResponse = Query1C("http://192.168.2.126/CHZMK/hs/Zakaz/DATA/" + project, "мобильный", "123456");
                obj = JObject.Parse(textResponse);
                msg = obj.SelectToken("Head").ToString();
                DrawingDataHead reflectionDataObject = JsonConvert.DeserializeObject<DrawingDataHead>(msg);

                if (reflectionDataObject.Code == -1)
                {
                    MessageBox.Show(reflectionDataObject.Message);
                    return null;
                }

                // Если данные по чертежам успешно получены               
                if (reflectionDataObject.Code == 0)
                {
                    double totalWeight = 0;
                    List<DateTime> alldates = new List<DateTime>();
                    Dates dates = new Dates();

                    drawingDataFrom1C.Table = obj.SelectToken("Table").ToObject<string[][]>();
                    for (int i = 0; i < drawingDataFrom1C.Table.Length; i++)
                    {
                        var item = drawingDataFrom1C.Table[i];
                        if (item[4].ToString() != "Полуфабрикат" | Convert.ToInt32(item[1].ToString()) == 0)
                        {
                            drawingFrom1CList.Add(new Drawing1C
                            {
                                DateOfEntry = item[3].ToString(),
                                IsXDraw = Convert.ToInt32(item[1].ToString()),
                                TypeOfNomenclature = item[4].ToString(),
                                Weight = item[2].ToString(),
                            });
                            totalWeight += System.Math.Round(Convert.ToDouble(item[2].ToString()), 1);

                            string[] splittedDT = item[3].ToString().Split('.');
                            DateTime DrawDate =
                                new DateTime(Convert.ToInt32(splittedDT[2]), Convert.ToInt32(splittedDT[1]), Convert.ToInt32(splittedDT[0]));

                            if (i == 0)
                            {
                                dates.MinDate = DrawDate;
                                dates.MaxDate = DrawDate;
                            }

                            alldates.Add(DrawDate);

                        }
                    }

                    totalWeight = System.Math.Round(totalWeight, 1);
                    ProjectsWeightDic.Add(project, totalWeight);

                    dates.MinDate = alldates.Min();
                    dates.MaxDate = alldates.Max();
                    ProjectsDates1C.Add(project, dates);
                }
                return drawingFrom1CList;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace.ToString());
                return null;
            }
        }
        public string Query1C(string adress, string username, string password)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(adress);

                string authInfo = username + ":" + password;
                authInfo = Convert.ToBase64String(Encoding.UTF8.GetBytes(authInfo));
                request.Headers["Authorization"] = "Basic " + authInfo;

                var response = (HttpWebResponse)request.GetResponse();
                using (var sreader = new StreamReader(response.GetResponseStream()))
                {
                    responseString = sreader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
            return responseString;
        }










        private void button42_Click(object sender, EventArgs e)
        {
            TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();
            try
            {
                if (drawingHandler.GetConnectionStatus())
                {
                    TSD.UI.Picker picker = drawingHandler.GetPicker();
                    TSD.PointList points = new TSD.PointList();

                    TSD.StringList prompts = new TSD.StringList
                    {
                        "Укажите первое положение",
                        "Укажите второе положение"
                    };

                    picker.PickPoints(2, prompts, out points, out TSD.ViewBase view);

                    TSD.Line Line = new TSD.Line(view, points[0], points[1]);
                    Line.Attributes.Line.Color = TSD.DrawingColors.Cyan;
                    Line.Insert();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            string path = @"SOFTWARE\Microsoft\NET Framework Setup\NDP";
            List<string> display_framwork_name = new List<string>();

            RegistryKey installed_versions = Registry.LocalMachine.OpenSubKey(path);
            string[] version_names = installed_versions.GetSubKeyNames();

            for (int i = 1; i <= version_names.Length - 1; i++)
            {
                string temp_name = "Microsoft .NET Framework " + version_names[i].ToString() + "  SP" + installed_versions.OpenSubKey(version_names[i]).GetValue("SP");
                display_framwork_name.Add(temp_name);
            }
        }

        private void button44_Click(object sender, EventArgs e)
        {
            string path = string.Empty;
            path = @"%AppData%\stuff";
            path = @"%aPpdAtA%\HelloWorld";
            path = @"%progRAMfiLES%\Adobe;%LOCALAPPDATA%\FileZilla"; // collection of paths
            path = @"%ProgramData%";

            path = Environment.ExpandEnvironmentVariables(path);
            var OS = Environment.OSVersion;
            string Version = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion", "ProductName", null);
            string releaseId = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId", "").ToString();

            // НУЖНЫ АДМИНСКИЕ ПРАВА ДЛЯ СОЗДАНИЯ SOURCE. Идея накрылась ...
            try
            {
                string source = "modelplugin";
                string log = "Application";
                if (!EventLog.SourceExists(source))
                {
                    EventLog.CreateEventSource(source, log);
                }
                EventLog.WriteEntry(source, "First message from the demo log within Application", EventLogEntryType.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //string log = "Windows PowerShell";
            //EventLog demoLog = new EventLog(log);
            //EventLogEntryCollection entries = demoLog.Entries;
            //foreach (EventLogEntry entry in entries)
            //{
            //    Console.WriteLine("Level: {0}", entry.EntryType);
            //    Console.WriteLine("Event id: {0}", entry.InstanceId);
            //    Console.WriteLine("Message: {0}", entry.Message);
            //    Console.WriteLine("Source: {0}", entry.Source);
            //    Console.WriteLine("Date: {0}", entry.TimeGenerated);
            //    Console.WriteLine("--------------------------------");
            //}

            //System.Diagnostics.EventLog appLog =new System.Diagnostics.EventLog();
            //appLog.Source = "This Application's Name";
            //appLog.WriteEntry("An entry to the Application event log.");

            //using (EventLog eventLog = new EventLog("Application"))
            //{
            //    eventLog.Source = "Application";
            //    eventLog.WriteEntry("Log message example", EventLogEntryType.Error);
            //}

            //// Create an instance of EventLog
            //System.Diagnostics.EventLog eventLog = new System.Diagnostics.EventLog();

            //var logs = EventLog.GetEventLogs();

            //EventLog.CreateEventSource("AAATestApplication", "Application");

            //// Check if the event source exists. If not create it.
            //if (!System.Diagnostics.EventLog.SourceExists("AAATestApplication"))
            //{
            //    System.Diagnostics.EventLog.CreateEventSource("AAATestApplication", "Application");
            //}

            //// Set the source name for writing log entries.
            //eventLog.Source = "TestApplication";

            //// Create an event ID to add to the event log
            //int eventID = 8;

            //// Write an entry to the event log.
            //eventLog.WriteEntry("test",
            //                    System.Diagnostics.EventLogEntryType.Error,
            //                    eventID);

            //// Close the Event Log
            //eventLog.Close();
        }

        private void button45_Click(object sender, EventArgs e)
        {
            string VersionDirectory = Tekla.Structures.Dialog.UIControls.EnvironmentVariables.GetEnvironmentVariable("XSDATADIR");
        }

        private void button46_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    TSD.DrawingEnumerator drawingEnumerator = new TSD.DrawingHandler().GetDrawingSelector().GetSelected();

            //    while (drawingEnumerator.MoveNext())
            //    {
            //        var selDraw = drawingEnumerator.Current;

            //        PropertyInfo pi = drawingEnumerator.Current.GetType()
            //            .GetProperty("Identifier", BindingFlags.Instance | BindingFlags.NonPublic);
            //        object value = pi.GetValue(drawingEnumerator.Current, null);
            //        Identifier Identifier = (Identifier)value;
            //        TSM.Beam fakebeam = new TSM.Beam();
            //        fakebeam.Identifier = Identifier;
            //        string drawType = "";
            //        fakebeam.GetReportProperty("TYPE", ref drawType);

            //        string oldZak = "";
            //        selDraw.GetUserProperty("metcon_Zakaz", ref oldZak);

            //        selDraw.SetUserProperty("metcon_DrZakaz", oldZak);
            //    }
            //}

            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}

            try
            {
                Dictionary<string, int> dict = new Dictionary<string, int>();
                string projResponse = DBQueryFree("select id, title from public.pkoservices_projects_book");
                JObject jObject = JObject.Parse(projResponse);
                var msg = jObject.SelectToken("msg");

                int projectID = int.MinValue;
                string title = string.Empty;
                if (msg.HasValues)
                {
                    foreach (var item in msg)
                    {
                        projectID = (int)item["id"];
                        title = (string)item["title"];

                        dict.Add(title, projectID);
                    }
                }

                TSD.DrawingEnumerator drawingEnumerator = new TSD.DrawingHandler().GetDrawingSelector().GetSelected();

                while (drawingEnumerator.MoveNext())
                {
                    var selDraw = drawingEnumerator.Current;

                    //string oldZak = "";
                    //selDraw.GetUserProperty("metcon_DrZakaz", ref oldZak);
                    //int projID = int.MinValue;
                    //dict.TryGetValue(oldZak, out projID);

                    PropertyInfo pi = drawingEnumerator.Current.GetType().GetProperty("Identifier", BindingFlags.Instance | BindingFlags.NonPublic);
                    object value = pi.GetValue(drawingEnumerator.Current, null);
                    Identifier Identifier = (Identifier)value;
                    TSM.Beam fakebeam = new TSM.Beam
                    {
                        Identifier = Identifier
                    };
                    string nameBase = "";
                    fakebeam.GetReportProperty("NAME_BASE", ref nameBase);

                    string drNum = string.Empty;
                    selDraw.GetUserProperty("metcon_KMD_Number", ref drNum);
                    drNum = drNum.Replace("и1", "");

                    string query = DBQueryFree(@"INSERT INTO public.model_drawing_enum(
                        model_id, type, name_base, list, listov, username, kmd_number, project_id)
	                    VALUES ('4aeffcfc-fa2c-46c6-8b2b-074e4be52569', 'A', '" + nameBase + "', '1', '1', 'skt', "
                        + Convert.ToInt32(drNum) + ", 60)");



                    //string drawResponse = DBQueryFree(@"
                    //        update public.model_drawing_enum
                    //        set kmd_number=" + Convert.ToInt32(drNum) + @"
                    //        where model_id='92de03bb-3c15-4d26-b6b8-f9142d85d128'
                    //        and name_base='" + nameBase + "'");
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void button47_Click(object sender, EventArgs e)
        {
            string[] rows = File.ReadAllLines(@"\\172.16.0.23\models\Temp\новый 1.txt");
            foreach (var item in rows)
            {
                string[] splitted = item.Split('\t');
                string query = DBQueryFree(@"INSERT INTO public.model_drawing_enum(
                        model_id, type, name_base, list, listov, username, kmd_number, project_id)
	                    VALUES ('" + splitted[1].ToString() + "', '" + splitted[2].ToString()
                    + "', '" + splitted[3].ToString() + "', '" + splitted[4].ToString()
                    + "', '" + splitted[5].ToString() + "', '" + splitted[6].ToString()
                    + "', " + Convert.ToInt32(splitted[9].ToString()) + ", " + Convert.ToInt32(splitted[10].ToString()) + ")");


                //if (splitted[10].ToString() == @"\N")
                //{
                //    long id = (long)Convert.ToDouble(splitted[0].ToString());
                //    string modelId= splitted[1].ToString();

                //    string query = DBQueryFree(@"
                //            update public.model_drawing_enum
                //            set model_id='" + modelId + "' where id=" + id);
                //}
            }
        }

        private void button48_Click(object sender, EventArgs e)
        {
            var drawings = new TSD.DrawingHandler().GetDrawings();

            drawings.SelectInstances = false;
            while (drawings.MoveNext())
            {
                //if (drawings.Current is TSD.AssemblyDrawing Dwg)
                //{
                var nsme = drawings.Current.Name;
                //}
            };

            //foreach (var item in drawings)
            //{
            //    TSD.Drawing drawing = (TSD.Drawing)item;
            //    var nsme = drawing.Name;
            //}
        }

        private void button49_Click(object sender, EventArgs e)
        {
            CultureInfo ci = Thread.CurrentThread.CurrentCulture;
            DayOfWeek fdow = ci.DateTimeFormat.FirstDayOfWeek;
            DayOfWeek today = DateTime.Now.DayOfWeek;
            DateTime firstDayNextWeek = DateTime.Now.AddDays(-(today - fdow)).Date.AddDays(7);
            DateTime lastDayNextWeek = firstDayNextWeek.AddDays(6);

            DateTime[] dateTimes = new DateTime[2]
            {
                firstDayNextWeek,
                lastDayNextWeek
            };
        }

        private void button50_Click(object sender, EventArgs e)
        {
            //https://csharp.webdelphi.ru/rabota-s-arxivami-zip-v-c/
            var githubToken = "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";

            var url = "https://github.com/skzlat/metcon_WPFPlugin_TableFromExcel/archive/refs/tags/v23.03.22.1.zip";
            var path = @"D:\TEMP\v23.03.22.1.zip";


            DirectoryInfo destination = new DirectoryInfo(@"D:\TEMP\DWN");
            if (!destination.Exists)
                destination.Create();
            using (var client = new System.Net.Http.HttpClient())
            {
                var credentials = string.Format(CultureInfo.InvariantCulture, "{0}:", githubToken);
                credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(credentials));
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", credentials);
                var contents = client.GetByteArrayAsync(url).Result;
                File.WriteAllBytes(path, contents);

                MemoryStream ms = new MemoryStream(contents);

                using (ZipArchive archive = new ZipArchive(ms))
                {
                    string fileName = "Functions.cs";
                    var file = archive.Entries.Where(s => s.Name == fileName).FirstOrDefault();
                    if (file != null)
                        file.ExtractToFile(Path.Combine(destination.FullName, fileName), true);

                    string directory = @"metcon_WPFPlugin_TableFromExcel-23.03.22.1\metcon_WPFPlugin_TableFromExcel\Properties";
                    var dirName = new DirectoryInfo(directory).Name;

                    var result = from currEntry in archive.Entries
                                 where Path.GetDirectoryName(currEntry.FullName) == directory
                                 where !string.IsNullOrEmpty(currEntry.Name)
                                 select currEntry;

                    DirectoryInfo destDir = new DirectoryInfo(Path.Combine(destination.FullName, dirName));
                    if (!destDir.Exists)
                        destDir.Create();
                    foreach (ZipArchiveEntry entry in result)
                        entry.ExtractToFile(Path.Combine(destDir.FullName, entry.Name), true);
                }
            }

            //string archivePath = @"c:\archive\archive_min.zip";
            //using var fileStream = File.Open(archivePath, FileMode.Open);
            //ZipArchive archive = new ZipArchive(fileStream);
            ////ИЛИ
            //ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update);
        }

        private void button51_Click(object sender, EventArgs e)
        {
            var token = "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";

            var client = new GitHubClient(new ProductHeaderValue("UpdateSettings"));
            var tokenAuth = new Credentials(token); // This can be a PAT or an OAuth token.
            client.Credentials = tokenAuth;

            var rel = client.Repository.Release.GetAll("skzlat", "metcon_WPFPlugin_TableFromExcel");
            var sss = rel.Result;

            var release = client.Repository.Release.GetLatest("skzlat", "metcon_WPFPlugin_TableFromExcel");
            var latest = release.Result;

            string downloadUrl = release.Result.Url;




            var releaseArchive = client.Repository.Release.Get("skzlat", "metcon_WPFPlugin_TableFromExcel", "23.03.22.1");
            var contents = releaseArchive.Result;


            WebClient webClient = new WebClient();
            webClient.Headers.Add("user-agent", "Anything");
            webClient.DownloadFileTaskAsync(new Uri(latest.ZipballUrl), @"D:\TEMP\23.03.22.1.zip");
        }
    }
}