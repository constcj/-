// (C) Copyright 2018 by Microsoft 
//
using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using AApplication = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Drawing;
using System.Windows.Media;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.Generic;
using System.Windows.Input;
using System.Windows.Forms;
using System.Linq;
//using Autodesk . AutoCAD . Interop;
using megbox = System.Windows.Forms.MessageBox;
// This line is not mandatory, but improves loading performances
[assembly: CommandClass(typeof(小工具.MyCommands))]

namespace 小工具
{

    // This class is instantiated by AutoCAD for each document when
    // a command is called by the user the first time in the context
    // of a given document. In other words, non static data in this class
    // is implicitly per-document!
    public class MyCommands
    {
        Window1 w1;
        // The CommandMethod attribute can be applied to any public  member 
        // function of any public class.
        // The function should take no arguments and return nothing.
        // If the method is an intance member then the enclosing class is 
        // intantiated for each document. If the member is a static member then
        // the enclosing class is NOT intantiated.
        //
        // NOTE: CommandMethod has overloads where you can provide helpid and
        // context menu.

        // Modal Command with localized name
        /// <summary>
        /// 自动编号
        /// </summary>
        [CommandMethod("AutoN")]
        public void MyCommand() // This method can have any name
        {
            // Put your command code here
            w1 = new Window1();

            w1.Button1.Click += Button1_Click;

            w1.ShowDialog();

        }

        /// <summary>
        /// 图号和编码对应表
        /// </summary>
        [CommandMethod("MyMap")]
        public void MyCommand1() // This method can have any name
        {
            try
            {

                // Put your command code here
                EXL ee = new EXL();

                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                string co = "";
                string tuh = "";
                using (doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        foreach (ObjectId item in br)
                        {
                            var re = tr.GetObject(item, OpenMode.ForWrite);
                            if (re is BlockReference)
                            {
                                BlockReference bre = (BlockReference)re;
                                AttributeCollection abu = bre.AttributeCollection;
                                foreach (ObjectId item1 in abu)
                                {
                                    AttributeReference at1 = (AttributeReference)tr.GetObject(item1, OpenMode.ForWrite);
                                    if (at1.Tag == "图纸名称") //  if (at1.Tag == "图纸编号")// && at1 . TextString == "" 
                                    {
                                        tuh = at1.TextString;

                                    }
                                    else if (at1.Tag == "图纸编号") //else if (at1.Tag == "模板编号")//&& at1 . TextString == "" 
                                    { co = at1.TextString; }
                                }
                                if (co != "")
                                {
                                    ee.SetV(co, tuh);
                                    co = "";
                                    tuh = "";
                                }
                            }
                        }
                        tr.Commit();
                    }
                }
                ee.Col();
            }
            catch (System.Exception)
            {

            }
        }
        void Button1_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                Boolean page = false;

                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                int i = 1;

                string suffix = w1.TextBox1.Text;
                if (suffix.Contains("-"))
                {
                    string[] suffixs = suffix.Split('-');

                    int count = suffixs.Length;

                    string sufmid = "";

                    if (suffixs[count - 1] != "")
                    {
                        for (int j = 0; j < count - 1; j++)
                        {
                            sufmid = sufmid + suffixs[j] + "-";
                        }

                        suffix = sufmid;

                        if (Regex.IsMatch(suffixs[count - 1], "\\d"))
                        {
                            i = Convert.ToInt32(suffixs[count - 1]);
                        }
                    }
                }
                else
                {
                    suffix = suffix + "-";
                }

                using (doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        foreach (ObjectId item in br)
                        {
                            var re = tr.GetObject(item, OpenMode.ForWrite);
                            if (re is BlockReference)
                            {
                                BlockReference bre = (BlockReference)re;
                                AttributeCollection abu = bre.AttributeCollection;
                                foreach (ObjectId item1 in abu)
                                {
                                    AttributeReference at1 = (AttributeReference)tr.GetObject(item1, OpenMode.ForWrite);
                                    if (at1.Tag == "图纸编号" && (at1.TextString == "" || at1.TextString == null))
                                    {
                                        at1.TextString = suffix + i;

                                        page = true;

                                        i++;
                                    }
                                    else if (at1.Tag == "项目名称" && (at1.TextString == "" || at1.TextString == null))
                                    {
                                        at1.TextString = w1.TextBox2.Text;
                                        if (Regex.IsMatch(w1.TextBox3.Text, "[01]\\.\\d+"))
                                        {
                                            double cr = Convert.ToDouble(w1.TextBox3.Text);
                                            at1.WidthFactor = cr;

                                        }
                                    }
                                    else if (at1.Tag == "第" && (at1.TextString == "" || at1.TextString == null))
                                    {
                                        if (page)
                                        {
                                            at1.TextString = (i - 1).ToString();
                                        }
                                        else
                                        {
                                            at1.TextString = i.ToString();
                                        }

                                        page = false;
                                    }
                                }
                            }

                            page = false;
                        }

                        foreach (ObjectId item in br)
                        {
                            var re = tr.GetObject(item, OpenMode.ForWrite);
                            if (re is BlockReference)
                            {
                                BlockReference bre = (BlockReference)re;
                                AttributeCollection abu = bre.AttributeCollection;
                                foreach (ObjectId item1 in abu)
                                {
                                    AttributeReference at1 = (AttributeReference)tr.GetObject(item1, OpenMode.ForWrite);
                                    if (at1.Tag == "总" && (at1.TextString == "" || at1.TextString == null))
                                    {

                                        at1.TextString = (i - 1).ToString();

                                    }
                                }
                            }

                        }

                        tr.Commit();
                    }
                }
                w1.Close();

            }
            catch (System.Exception)
            {

                throw;
            }
        }
        /// <summary>
        /// 清楚图框内容
        /// </summary>
        [CommandMethod("CleM")]
        public void MyCommand4()
        {
            try
            {
                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                using (doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        foreach (ObjectId item in br)
                        {
                            var re = tr.GetObject(item, OpenMode.ForWrite);
                            if (re is BlockReference)
                            {
                                BlockReference bre = (BlockReference)re;
                                AttributeCollection abu = bre.AttributeCollection;
                                foreach (ObjectId item1 in abu)
                                {
                                    AttributeReference at1 = (AttributeReference)tr.GetObject(item1, OpenMode.ForWrite);
                                    if (at1.Tag == "图纸编号" && !Regex.IsMatch(at1.TextString, "^ZW-"))
                                    {
                                        at1.TextString = "";

                                    }
                                    else if (at1.Tag == "项目名称")
                                    {
                                        at1.TextString = "";

                                    }
                                    else if (at1.Tag == "总")
                                    {
                                        at1.TextString = "";
                                    }
                                    else if (at1.Tag == "第")
                                    {
                                        at1.TextString = "";
                                    }
                                }
                            }
                        }
                        tr.Commit();
                    }
                }

            }
            catch (System.Exception)
            {

                throw;
            }
        }
        private Window2 w2 = null;
        System.Windows.Media.Imaging.BitmapImage bit = null;
        /// <summary>
        /// 大样图图片
        /// </summary>
        [CommandMethod("MyPic")]
        public void MyCommand2() // This method can have any name
        {
            Document doc = AApplication.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            w2 = new Window2();
            string code = "";
            w2.KeyDown += w2_KeyDown;

            using (doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    PromptEntityResult sel = ed.GetEntity("选择模板编号");

                    if (sel.Status == PromptStatus.OK)
                    {
                        Entity en = (Entity)tr.GetObject(sel.ObjectId, OpenMode.ForRead);
                        if (en is DBText)
                        {
                            DBText dtext = (DBText)en;
                            code = dtext.TextString;
                        }


                    }
                    tr.Commit();
                }
            }

            bit = GetBit(code);

            if (bit != null)
            {
                w2.Image.Source = bit;

                w2.ShowDialog();
            }

        }

        void w2_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                bit = null;
                w2.Close();
            };
        }

        private System.Windows.Media.Imaging.BitmapImage GetBit(string code)
        {
            System.Windows.Media.Imaging.BitmapImage bi = null;
            try
            {
                DirectoryInfo dir = new DirectoryInfo("F:\\大样图文件库");
                FileInfo[] fis = dir.GetFiles("*.png");

                Dictionary<string, string> dd = new Dictionary<string, string>();

                foreach (FileInfo item in fis)
                {
                    string frma = Regex.Replace(item.Name, ".png", replacement: "");
                    dd.Add(frma, item.FullName);
                }
                string codef = code;
                codef = Regex.Replace(codef, "/", replacement: "!");
                if (!dd.ContainsKey(codef))
                {
                    codef = Regex.Replace(codef, "\\d{1,4}\\.?\\d{0,4}|\\s", replacement: "");
                }

                if (dd.ContainsKey(codef))
                {
                    //  Document doc1 = AApplication . DocumentManager . MdiActiveDocument;
                    // Document doc = AApplication . DocumentManager . Open (dd[codef] , true );

                    //  doc . Window . Visible = false;
                    //  Database db =doc . Database;

                    //   using ( Database da = new Database ( false , true ) )
                    //  {
                    //   da . ReadDwgFile ( "F:\\大样图文件库\\"+ "Drawing1.dwg" , FileOpenMode . OpenForReadAndAllShare , false , null );

                    //  };
                    // using ( Transaction tr = db . TransactionManager . StartTransaction ( ) )
                    // {
                    //     BlockTable bt = ( BlockTable ) tr . GetObject ( db . BlockTableId , OpenMode . ForRead );
                    //   BlockTableRecord br = ( BlockTableRecord ) tr . GetObject ( bt [ BlockTableRecord . ModelSpace ] , OpenMode . ForRead );                                     
                    //    foreach ( var item in br )
                    //  {
                    //       var en= tr . GetObject ( item , OpenMode . ForRead );
                    //    if ( en is BlockReference )
                    //       {
                    //       BlockReference bre=( BlockReference ) en;
                    // uint wit=Convert . ToUInt32 ( bre . GeometricExtents . MaxPoint . Y );
                    //  uint len=Convert . ToUInt32 ( bre . GeometricExtents . MaxPoint . X );

                    //  bi =new Bitmap (Convert.ToInt32( w2.Width), Convert.ToInt32(w2.Height));
                    //  bi = doc . CapturePreviewImage ( len ,wit);
                    //          Bitmap  bi = doc . CapturePreviewImage ( Convert . ToUInt32 ( bre.GeometricExtents.MaxPoint.X ) , Convert . ToUInt32 ( bre.GeometricExtents.MaxPoint.Y  ) );                              

                    //w2 . Image . Source = System . Windows . Interop . Imaging . CreateBitmapSourceFromHBitmap ( bi . GetHbitmap ( ) , IntPtr . Zero , Int32Rect . Empty , BitmapSizeOptions . FromEmptyOptions ( ) );
                    bi = new BitmapImage(new Uri(dd[codef], UriKind.Absolute));
                    //     break;
                    //  }
                    // }
                    //  tr . Commit ( );
                    //  doc . CloseAndDiscard ( );
                    //   }

                }
            }
            catch (System.Exception)
            {

            }
            return bi;
        }
        /// <summary>
        /// 大样图分开存放
        /// </summary>
        [CommandMethod("CIP")]
        public void MyCommand5() // This method can have any name
        {
            try
            {

                OpenFileDialog openDialog = new OpenFileDialog();

                // openDialog . Filter = "Excel File(*.xlsx)|*.xlsx";

                SortedSet<string> sedaring = new SortedSet<string>();

                string filna = "图纸";

                if (DialogResult.OK == openDialog.ShowDialog())
                {
                    string filename = openDialog.FileName;

                    filna = Path.ChangeExtension(filename, "dwg");

                    EXLe exl = new EXLe(filename);

                    sedaring = exl.GetDar();

                }
                else
                {
                    return;
                }

                if (sedaring.Count == 0)
                {
                    return;
                }

                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                //   Point3d loca=new Point3d ( );
                DBObjectCollection dc = new DBObjectCollection();

                ObjectIdCollection oj = new ObjectIdCollection();

                Dictionary<string, BlockReference> mabl = new Dictionary<string, BlockReference>();

                using (doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {

                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                        foreach (ObjectId item in br)
                        {
                            var re = tr.GetObject(item, OpenMode.ForRead);
                            if (re is BlockReference)
                            {
                                BlockReference bre = (BlockReference)re.Clone();
                                AttributeCollection abu = bre.AttributeCollection;
                                foreach (AttributeReference item1 in abu)
                                {
                                    AttributeReference at1 = item1;
                                    if (at1.Tag == "图纸编号")// && at1 . TextString == "" 
                                    {
                                        if (at1.TextString != null && at1.TextString != "")
                                        {
                                            if (!mabl.ContainsKey(at1.TextString))
                                            {
                                                mabl.Add(at1.TextString, bre);
                                            }
                                            else
                                            { megbox.Show(at1.TextString + "图纸重复存在"); }

                                        }
                                    }
                                }
                            }
                        }

                        tr.Commit();
                    }
                }

                using (doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {

                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        foreach (string tuh in sedaring)
                        {
                            if (mabl.ContainsKey(tuh))
                            {
                                BlockReference bre = mabl[tuh];
                                AttributeCollection abu = bre.AttributeCollection;

                                Point3d minp = new Point3d(bre.GeometricExtents.MinPoint.X, bre.GeometricExtents.MinPoint.Y, 0);
                                Point3d maxp = new Point3d(bre.GeometricExtents.MaxPoint.X, bre.GeometricExtents.MaxPoint.Y, 0);

                                PromptSelectionResult prs = ed.SelectCrossingWindow(minp, maxp);

                                if (prs.Status == PromptStatus.OK)
                                {
                                    SelectionSet sse = prs.Value;

                                    foreach (var item2 in sse.GetObjectIds())
                                    {
                                        oj.Add(item2);
                                    }
                                }
                            }

                        }

                        tr.Commit();
                    }

                    if (oj.Count != 0)
                    {
                        IdMapping i1 = new IdMapping();
                        using (Database da = new Database(false, true))
                        {
                            da.ReadDwgFile(MyPlugin.direct + "\\Drawing1.dwg", FileOpenMode.OpenForReadAndAllShare, false, null);

                            using (Transaction tr1 = da.TransactionManager.StartTransaction())
                            {
                                BlockTable bt1 = (BlockTable)tr1.GetObject(da.BlockTableId, OpenMode.ForRead);
                                BlockTableRecord br1 = (BlockTableRecord)tr1.GetObject(bt1[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                                da.WblockCloneObjects(oj, da.CurrentSpaceId, i1, DuplicateRecordCloning.Ignore, false);
                                tr1.Commit();

                            }
                            da.SaveAs(filna, DwgVersion.AC1027);
                        }

                    }
                }

            }
            catch (System.Exception)
            {
            }
        }
        ///// <summary>
        ///// 从数据大样图提取
        ///// </summary>
        //[CommandMethod ( "CTIP" )]
        //public void MyCommand7 ( ) // This method can have any name
        //{
        //    OpenFileDialog openDialog = new OpenFileDialog ( );

        //    openDialog . Filter = "Excel File(*.xlsx)|*.xlsx";
        //    openDialog . Multiselect = true;
        //    SortedSet<string> sedaring=new SortedSet<string> ( );

        //    string filna="图纸.dwg";

        //    if ( DialogResult . OK == openDialog . ShowDialog ( ) )
        //    {
        //        string[] filename = openDialog . FileNames;



        //        filna = Path . ChangeExtension ( filename , "dwg" );

        //        if ( filna . Contains ( "清单" ) )
        //        {
        //            filna = filna . Replace ( "清单" , "" );
        //        }

        //        EXLe exl=new EXLe ( filename );

        //        sedaring = exl . GetDar1 ( );

        //    }
        //    else
        //    {
        //        return;
        //    }

        //    if ( sedaring . Count == 0 )
        //    {
        //        return;
        //    }

        //    Document doc = AApplication . DocumentManager . MdiActiveDocument;

        //    Editor ed = doc . Editor;

        //    Database db = doc . Database;

        //    //   Point3d loca=new Point3d ( );
        //    DBObjectCollection dc=new DBObjectCollection ( );

        //    ObjectIdCollection oj=new ObjectIdCollection ( );

        //    Dictionary<string,BlockReference> mabl=new Dictionary<string , BlockReference> ( );

        //    using ( doc . LockDocument ( ) )
        //    {
        //        using ( Transaction tr = db . TransactionManager . StartTransaction ( ) )
        //        {

        //            BlockTable bt = ( BlockTable ) tr . GetObject ( db . BlockTableId , OpenMode . ForRead );
        //            BlockTableRecord br = ( BlockTableRecord ) tr . GetObject ( bt [ BlockTableRecord . ModelSpace ] , OpenMode . ForRead );

        //            foreach ( ObjectId item in br )
        //            {
        //                var re=tr . GetObject ( item , OpenMode . ForRead );
        //                if ( re is BlockReference )
        //                {
        //                    BlockReference bre=( BlockReference ) re . Clone ( );
        //                    AttributeCollection abu=bre . AttributeCollection;
        //                    foreach ( AttributeReference item1 in abu )
        //                    {
        //                        AttributeReference at1=item1;
        //                        if ( at1 . Tag == "模板编号" )// && at1 . TextString == "" 
        //                        {                               
        //                            if ( at1 . TextString != null && at1 . TextString != "" )
        //                            {
        //                                string cod=Regex.Replace( at1 . TextString,pattern:"\\([A-Z]+\\)|\\s",replacement:"");

        //                                if ( !mabl . ContainsKey (cod) )
        //                                {
        //                                    mabl . Add ( at1 . TextString , bre );
        //                                }
        //                                else
        //                                { 
        //                                    megbox . Show ( at1 . TextString + "图纸重复存在" ); }

        //                            }
        //                        }
        //                    }
        //                }
        //            }

        //            tr . Commit ( );
        //        }
        //    }

        //    using ( doc . LockDocument ( ) )
        //    {
        //        using ( Transaction tr = db . TransactionManager . StartTransaction ( ) )
        //        {

        //            BlockTable bt = ( BlockTable ) tr . GetObject ( db . BlockTableId , OpenMode . ForRead );
        //            BlockTableRecord br = ( BlockTableRecord ) tr . GetObject ( bt [ BlockTableRecord . ModelSpace ] , OpenMode . ForWrite );

        //            foreach ( string tuh in sedaring )
        //            {
        //                string tuh1=tuh;

        //                if ( !mabl.ContainsKey(tuh1) )
        //                {
        //                    tuh1 = Regex . Replace ( tuh1 , pattern: "(?<!\\()\\d{1,4}|\\s" , replacement: "" );
        //                }

        //                if ( mabl . ContainsKey ( tuh1 ) )
        //                {
        //                    BlockReference bre=mabl [ tuh1 ];
        //                    AttributeCollection abu=bre . AttributeCollection;

        //                    Point3d minp= new Point3d ( bre . GeometricExtents . MinPoint . X , bre . GeometricExtents . MinPoint . Y , 0 );
        //                    Point3d maxp= new Point3d ( bre . GeometricExtents . MaxPoint . X , bre . GeometricExtents . MaxPoint . Y , 0 );

        //                    PromptSelectionResult prs=ed . SelectCrossingWindow ( minp , maxp );

        //                    if ( prs . Status == PromptStatus . OK )
        //                    {
        //                        SelectionSet sse=prs . Value;

        //                        foreach ( var item2 in sse . GetObjectIds ( ) )
        //                        {
        //                            oj . Add ( item2 );
        //                        }
        //                    }
        //                }

        //            }

        //            tr . Commit ( );
        //        }

        //        if ( oj . Count != 0 )
        //        {
        //            IdMapping i1=new IdMapping ( );
        //            using ( Database da = new Database ( false , true ) )
        //            {
        //                da . ReadDwgFile ( MyPlugin . direct + "\\Drawing1.dwg" , FileOpenMode . OpenForReadAndAllShare , false , null );

        //                using ( Transaction tr1 = da . TransactionManager . StartTransaction ( ) )
        //                {
        //                    BlockTable bt1 = ( BlockTable ) tr1 . GetObject ( da . BlockTableId , OpenMode . ForRead );
        //                    BlockTableRecord br1 = ( BlockTableRecord ) tr1 . GetObject ( bt1 [ BlockTableRecord . ModelSpace ] , OpenMode . ForWrite );

        //                    da . WblockCloneObjects ( oj , da . CurrentSpaceId , i1 , DuplicateRecordCloning . Ignore , false );
        //                    tr1 . Commit ( );

        //                }
        //                da . SaveAs ( filna , DwgVersion . AC1027 );
        //            }

        //        }
        //    }
        //}
        /// <summary>
        /// 检查
        /// </summary>
        [CommandMethod("CNU")]
        public void MyCommand6() // This method can have any name
        {

            Boolean hav = false;

            EXLg eX = new EXLg(ref hav);
            try
            {

                if (hav == true)
                {

                    Document doc = AApplication.DocumentManager.MdiActiveDocument;

                    Editor ed = doc.Editor;

                    Database db = doc.Database;

                    //   Point3d loca=new Point3d ( );
                    DBObjectCollection dc = new DBObjectCollection();

                    ObjectIdCollection oj = new ObjectIdCollection();

                    Dictionary<string, BlockReference> mabl = new Dictionary<string, BlockReference>();
                    Dictionary<string, int> codes = new Dictionary<string, int>();

                    Dictionary<string, string> mapcode = new Dictionary<string, string>();

                    string str = File.ReadAllText(MyPlugin.direct + "\\code.txt");

                    string strcode1 = File.ReadAllText(MyPlugin.direct + "\\code1.txt");
                    string[] strcode1s = strcode1.Split('!');

                    str = str.Replace("\r\n", "");

                    string[] strs = str.Split(';');

                    foreach (string item in strs)
                    {
                        string[] strs1 = item.Split(':');

                        if (!mapcode.ContainsKey(strs1[0]))
                        {
                            mapcode.Add(strs1[0], strs1[1]);

                        }

                    }

                    using (doc.LockDocument())
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                            foreach (var item in br)
                            {
                                var dtex = tr.GetObject(item, OpenMode.ForRead);
                                string code1 = "";

                                string suf = "";

                                if (dtex is DBText)
                                {
                                    DBText dBText = (DBText)dtex;
                                    code1 = dBText.TextString;

                                }
                                else if (dtex is MText)
                                {
                                    MText dBText = (MText)dtex;
                                    code1 = dBText.Text;
                                }
                                else
                                {
                                    continue;
                                }

                                if (code1 != "")
                                {

                                    code1 = code1.Replace(" ", "");

                                    if (Regex.IsMatch(code1, "^([MLQDTJ]|[MLQDTJ]B\\d)-"))
                                    {
                                        suf = Regex.Match(code1, "^([MLQDTJ]|[MLQDTJ]B\\d)-").Value;
                                        code1 = Regex.Replace(code1, "^([MLQDTJ]|[MLQDTJ]B\\d)-", "");
                                    }
                                    if (Regex.IsMatch(code1, "\\(?[\u4e00-\u9fa5]+\\)?")) code1 = Regex.Replace(code1, "\\(?[\u4e00-\u9fa5]+\\)?", "");

                                    if (code1 == "") continue;

                                    string lcoded = "";

                                    for (int i = 0; i < strcode1s.Length; i++)
                                    {
                                        lcoded = Regex.Match(code1, strcode1s[i]).Value;

                                        if (lcoded != "") break;

                                    }

                                    lcoded = Regex.Replace(lcoded, "\\d{1,3}\\.?\\d{0,3}", "#");

                                    if (lcoded != "" && mapcode.ContainsKey(lcoded))
                                    {

                                        if (Regex.IsMatch(code1, mapcode[lcoded]))
                                        {
                                            string code3 = code1;
                                            int nub = 1;
                                            if (Regex.IsMatch(code1, "(\\(\\d\\)|\\[\\d\\])$"))
                                            {
                                                string codenub = Regex.Match(code1, "(\\(\\d\\)|\\[\\d\\])$").Value;
                                                nub = Convert.ToInt32(Regex.Match(codenub, "\\d").Value);
                                                code3 = Regex.Replace(code1, "(\\(\\d\\)|\\[\\d\\])$", "");
                                            }

                                            if (suf != "") code3 = suf + code3;

                                            if (codes.ContainsKey(code3))
                                            {
                                                codes[code3] = codes[code3] + nub;
                                            }
                                            else
                                            {
                                                codes[code3] = nub;
                                            }
                                        }

                                    }

                                }

                            }
                            tr.Commit();
                        }
                    }

                    eX.setValueRange(codes);

                }

                megbox.Show("计算完成！");

            }

            catch (System.Exception E)
            {

                MyPlugin.ExceptionWrit(E);
            }
        }

        /// <summary>
        /// 检查文字重叠
        /// </summary>
        [CommandMethod("CDB")]
        public void MyCommand8() // This method can have any name
        {

            try
            {

                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                //   Point3d loca=new Point3d ( );
                //DBObjectCollection dc = new DBObjectCollection();
                List<DBText> dBL1 = new List<DBText>();

                // List<CRegion> plist = new List<CRegion>();
                //  List<Point3dCollection> plist = new List<Point3dCollection>();
                // List<DBText> plistd = new List<DBText>();

                Dictionary<Point3dCollection, DBText> pd = new Dictionary<Point3dCollection, DBText>();

                TypedValue[] tl = new TypedValue[] { new TypedValue((int)DxfCode.Operator, "<or"), new TypedValue(0, "TEXT"), new TypedValue((int)DxfCode.Operator, "or>") };

                // TypedValue[] tl = new TypedValue[] {
                //  new TypedValue((int)DxfCode.Operator, "<and"), new TypedValue(0, "POLYLINE"), new TypedValue((int)DxfCode.Operator, "and>")};

                SelectionFilter fu = new SelectionFilter(tl);
                PromptSelectionOptions selop = new PromptSelectionOptions();
                PromptSelectionResult sel = ed.GetSelection(selop, fu);

                using (doc.LockDocument())
                {
                    if (sel.Status == PromptStatus.OK)
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                            BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                            SelectionSet seva = sel.Value;
                            foreach (var item in seva.GetObjectIds())
                            {
                                Entity dtex = (Entity)tr.GetObject(item, OpenMode.ForRead);

                                if (dtex is DBText)
                                {

                                    DBText dBText = (DBText)dtex;

                                    dBL1.Add(dBText);

                                    Point3dCollection point3DCollection = getDBp(dBText, 0);

                                    pd.Add(point3DCollection, dBText);

                                }
                                else

                                {
                                    continue;
                                }
                            }

                            while (dBL1.Count > 0)
                            {
                                PromptSelectionResult prs = ed.SelectCrossingPolygon(pd.Keys.ElementAt(0), fu);

                                if (prs.Status == PromptStatus.OK)
                                {
                                    SelectionSet sse = prs.Value;

                                    foreach (var item in sse.GetObjectIds())
                                    {
                                        Entity dtexM = (Entity)tr.GetObject(item, OpenMode.ForRead);
                                        if (dtexM is DBText)
                                        {
                                            DBText dB11 = dtexM as DBText;
                                            dBL1.Remove(dB11);
                                        }
                                    }
                                    if (sse.Count > 1)
                                    {
                                        Point3d ps = pd.Keys.ElementAt(0)[0];
                                        Line line = new Line(new Point3d(0, 0, 0), ps);
                                        br.AppendEntity(line);
                                        tr.AddNewlyCreatedDBObject(line, true);
                                    }
                                }
                                else
                                {
                                    dBL1.Remove(pd[pd.Keys.ElementAt(0)]);
                                }
                                pd.Remove(pd.Keys.ElementAt(0));
                            }
                            //    //foreach (Polyline3d item in polylines)
                            //    //{
                            //    //    plist.Remove(item);

                            //    //    Point3dCollection point = new Point3dCollection();
                            //    //    if (plist.Count > 0)
                            //    //    {
                            //    //        foreach (Polyline3d item1 in plist)
                            //    //        {
                            //    //            // CRegion cRegionitem = item.Clone() as CRegion;
                            //    //            //CRegion cRegionitem1 = item1.Clone() as CRegion;
                            //    //            // cRegionitem.BooleanOperation(BooleanOperationType.BoolIntersect, cRegionitem1);

                            //    //            item.IntersectWith(item1, Intersect.OnBothOperands, point, 0, 0);

                            //    //            //  if (cRegionitem.Area > 10)
                            //    //            if (point.Count > 0)
                            //    //            {
                            //    //                Point3d ps = point[0];

                            //    //                Line line = new Line(new Point3d(0, 0, 0), ps);
                            //    //                br.AppendEntity(line);
                            //    //                tr.AddNewlyCreatedDBObject(line, true);

                            //    //            }

                            //    //        }
                            //    //    }
                            //    //}
                            //}
                            tr.Commit();
                        }
                    }
                    megbox.Show("检查完成！");

                }
            }

            catch (System.Exception E)
            {

                MyPlugin.ExceptionWrit(E);
            }
        }

        /// <summary>
        /// 检查文字重叠
        /// </summary>
        [CommandMethod("CDBd")]
        public void MyCommand81() // This method can have any name
        {

            try
            {

                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                PromptDoubleResult result = ed.GetDouble("输入文字重叠边界");

                if (result.Status == PromptStatus.OK)
                {

                    double r = result.Value;

                    if (r > 0 && r <= 1)
                    {
                        Database db = doc.Database;

                        List<DBText> dBL1 = new List<DBText>();

                        Dictionary<Point3dCollection, DBText> pd = new Dictionary<Point3dCollection, DBText>();

                        TypedValue[] tl = new TypedValue[] { new TypedValue((int)DxfCode.Operator, "<or"), new TypedValue(0, "TEXT"), new TypedValue((int)DxfCode.Operator, "or>") };

                        SelectionFilter fu = new SelectionFilter(tl);
                        PromptSelectionOptions selop = new PromptSelectionOptions();
                        PromptSelectionResult sel = ed.GetSelection(selop, fu);

                        using (doc.LockDocument())
                        {
                            if (sel.Status == PromptStatus.OK)
                            {
                                using (Transaction tr = db.TransactionManager.StartTransaction())
                                {
                                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                                    BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                                    SelectionSet seva = sel.Value;
                                    foreach (var item in seva.GetObjectIds())
                                    {
                                        Entity dtex = (Entity)tr.GetObject(item, OpenMode.ForRead);

                                        if (dtex is DBText)
                                        {

                                            DBText dBText = (DBText)dtex;

                                            dBL1.Add(dBText);

                                            Point3dCollection point3DCollection = getDBp(dBText, r);

                                            pd.Add(point3DCollection, dBText);

                                        }
                                        else

                                        {
                                            continue;
                                        }
                                    }

                                    while (dBL1.Count > 0)
                                    {
                                        PromptSelectionResult prs = ed.SelectCrossingPolygon(pd.Keys.ElementAt(0), fu);

                                        if (prs.Status == PromptStatus.OK)
                                        {
                                            SelectionSet sse = prs.Value;

                                            foreach (var item in sse.GetObjectIds())
                                            {
                                                Entity dtexM = (Entity)tr.GetObject(item, OpenMode.ForRead);
                                                if (dtexM is DBText)
                                                {
                                                    DBText dB11 = dtexM as DBText;
                                                    dBL1.Remove(dB11);
                                                }
                                            }
                                            if (sse.Count > 1)
                                            {
                                                Point3d ps = pd.Keys.ElementAt(0)[0];
                                                Line line = new Line(new Point3d(0, 0, 0), ps);
                                                br.AppendEntity(line);
                                                tr.AddNewlyCreatedDBObject(line, true);
                                            }
                                        }
                                        else
                                        {
                                            dBL1.Remove(pd[pd.Keys.ElementAt(0)]);
                                        }
                                        pd.Remove(pd.Keys.ElementAt(0));
                                    }

                                    tr.Commit();
                                }
                            }
                            megbox.Show("检查完成！");

                        }
                    }

                }

            }

            catch (System.Exception E)
            {

                MyPlugin.ExceptionWrit(E);
            }
        }

        [CommandMethod("CFQ")]
        public void MyCommand10()
        {
            try
            {
                List<string> layname = new List<string>();
                // string layname = "";
                string laynamemid = "";

                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                Dictionary<string, Dictionary<string, int>> mabl = new Dictionary<string, Dictionary<string, int>>();

                Dictionary<string, string> mapcode = new Dictionary<string, string>();

                string str = File.ReadAllText(MyPlugin.direct + "\\code.txt");

                string strcode1 = File.ReadAllText(MyPlugin.direct + "\\code1.txt");
                string[] strcode1s = strcode1.Split('!');

                str = str.Replace("\r\n", "");

                string[] strs = str.Split(';');

                foreach (string item in strs)
                {
                    string[] strs1 = item.Split(':');

                    if (!mapcode.ContainsKey(strs1[0]))
                    {
                        mapcode.Add(strs1[0], strs1[1]);

                    }

                }

                List<List<Line>> plist = new List<List<Line>>();

                List<DBText> plistd = new List<DBText>();

                using (doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        LayerTable layeys = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

                        foreach (var item in layeys)
                        {
                            LayerTableRecord layeyscord = tr.GetObject(item, OpenMode.ForRead) as LayerTableRecord;
                            if (Regex.IsMatch(layeyscord.Name, "边框$"))
                            {
                                layname.Add(layeyscord.Name);
                                // layname = layeyscord.Name;
                            }
                        }

                        tr.Commit();
                    }

                    bool res = false;

                    //if (layname!="")
                    if (layname.Count > 0)
                    {
                        TypedValue[] t2 = new TypedValue[7 * layname.Count];

                        int st = 0;

                        for (int i = 0; i < layname.Count; i++)
                        {
                            st = 7 * i;

                            t2[st] = new TypedValue((int)DxfCode.Operator, "<and");
                            t2[st + 1] = new TypedValue((int)DxfCode.Operator, "<or");

                            t2[st + 2] = new TypedValue(0, "POLYLINE");
                            t2[st + 3] = new TypedValue(0, "LWPOLYLINE");

                            t2[st + 4] = new TypedValue((int)DxfCode.Operator, "or>");

                            t2[st + 5] = new TypedValue((int)DxfCode.LayerName, layname[i]);

                            t2[st + 6] = new TypedValue((int)DxfCode.Operator, "and>");

                        }

                        TypedValue[] tl = new TypedValue[3 + t2.Length];

                        tl[0] = new TypedValue((int)DxfCode.Operator, "<or");

                        for (int i = 0; i < t2.Length; i++)
                        {
                            tl[i + 1] = t2[i];
                        }

                        tl[tl.Length - 2] = new TypedValue(0, "TEXT");
                        tl[tl.Length - 1] = new TypedValue((int)DxfCode.Operator, "or>");

                        SelectionFilter fu = new SelectionFilter(tl);
                        PromptSelectionOptions selop = new PromptSelectionOptions();
                        PromptSelectionResult sel = ed.GetSelection(selop, fu);
                        DBObjectCollection dBObject = new DBObjectCollection();

                        if (sel.Status == PromptStatus.OK)
                        {
                            using (Transaction tr = db.TransactionManager.StartTransaction())
                            {
                                SelectionSet seva = sel.Value;

                                foreach (var item in seva.GetObjectIds())
                                {
                                    var dtexM = tr.GetObject(item, OpenMode.ForWrite);

                                    if (dtexM is Polyline3d)
                                    {

                                        Polyline3d dtex = (Polyline3d)tr.GetObject(item, OpenMode.ForRead);

                                        Point3d p1 = dtex.StartPoint;
                                        Point3d p2 = dtex.EndPoint;

                                        if (Math.Round(p1.X - p2.X, 5) != 0 || Math.Round(p1.Y - p2.Y, 5) != 0)
                                        {
                                            dtex.Closed = true;
                                        }

                                        dtex.Explode(dBObject);

                                        List<Line> lines = new List<Line>();

                                        foreach (var item1 in dBObject)
                                        {
                                            if (item1 is Line)
                                            {
                                                Line line = item1 as Line;

                                                if (!lines.Contains(line)) lines.Add(line);

                                            }
                                        }

                                        plist.Add(lines);

                                        laynamemid = dtex.Layer;

                                    }
                                    else if (dtexM is Polyline)
                                    {
                                        Polyline dtex = (Polyline)tr.GetObject(item, OpenMode.ForRead);

                                        List<Line> lines = new List<Line>();

                                        Point3d p1 = dtex.StartPoint;
                                        Point3d p2 = dtex.EndPoint;

                                        if (Math.Round(p1.X - p2.X, 5) != 0 || Math.Round(p1.Y - p2.Y, 5) != 0)
                                        {
                                            dtex.Closed = true;
                                        }

                                        dtex.Explode(dBObject);

                                        foreach (var item1 in dBObject)
                                        {
                                            if (item1 is Line)
                                            {
                                                Line line = item1 as Line;

                                                if (!lines.Contains(line)) lines.Add(line);

                                            }
                                        }

                                        plist.Add(lines);

                                        laynamemid = dtex.Layer;

                                    }
                                    else if (dtexM is DBText)
                                    {

                                        DBText dBText = (DBText)dtexM;
                                        // string code1 = dBText.TextString;

                                        plistd.Add(dBText);

                                        //  plistd.Add(new Point3d((dBText.GeometricExtents.MaxPoint.X + dBText.GeometricExtents.MinPoint.X) / 2,
                                        //   (dBText.GeometricExtents.MaxPoint.Y + dBText.GeometricExtents.MinPoint.Y) / 2, 0), code1);

                                    }

                                }
                                tr.Commit();
                            }

                            List<DBText> plistdcopy = new List<DBText>(plistd);

                            using (Transaction tr = db.TransactionManager.StartTransaction())
                            {
                                foreach (List<Line> item in plist)
                                {

                                    // List<double> xList = item.Select(x => x.X).ToList();
                                    // List<double> yList = item.Select(y => y.Y).ToList();
                                    string bo = "";
                                    Dictionary<string, int> codes = new Dictionary<string, int>();
                                    foreach (DBText dBText in plistd)
                                    {
                                        Point3d item11 = new Point3d((dBText.GeometricExtents.MaxPoint.X + dBText.GeometricExtents.MinPoint.X) / 2,
                                            (dBText.GeometricExtents.MaxPoint.Y + dBText.GeometricExtents.MinPoint.Y) / 2, 0);

                                        string code1 = dBText.TextString;

                                        if (PositionPnpoly(item, item11.X, item11.Y))
                                        {

                                            plistdcopy.Remove(dBText);

                                            if (!Regex.IsMatch(laynamemid, "^(吊模|墙)边框$"))
                                            {
                                                if (Regex.IsMatch(code1, "^[ABCDEFGH][QLM](1\\d{2}|\\d{1,2})$"))
                                                {
                                                    bo = Regex.Match(code1, "^[ABCDEFGH][QLM](1\\d{2}|\\d{1,2})$").Value;
                                                    continue;
                                                }
                                                else if (Regex.IsMatch(code1, "^节点J\\d{1,2}$"))
                                                {
                                                    bo = Regex.Match(code1, "^节点J\\d{1,2}$").Value;
                                                    continue;
                                                }
                                            }
                                            else
                                            {
                                                if (Regex.IsMatch(code1, "^[ABCDEFGH]$"))
                                                {
                                                    string lay = Regex.Match(laynamemid, "(吊模|墙)").Value;

                                                    bo = Regex.Match(code1, "^[ABCDEFGH]$").Value;
                                                    bo = lay + bo + "区";
                                                    continue;
                                                }
                                            }

                                            Dictionary<string, int> codes1 = getDBm(code1, mapcode, strcode1s, "(\\(2\\))$");

                                            if (codes1.Count > 0)
                                            {
                                                if (codes.ContainsKey(codes1.Keys.ElementAt(0)))
                                                {
                                                    codes[codes1.Keys.ElementAt(0)] = codes[codes1.Keys.ElementAt(0)] + codes1[codes1.Keys.ElementAt(0)];
                                                }
                                                else
                                                {
                                                    codes.Add(codes1.Keys.ElementAt(0), codes1[codes1.Keys.ElementAt(0)]);
                                                }
                                            }

                                        }
                                    }

                                    plistd = new List<DBText>(plistdcopy);

                                    if (bo != "")
                                    {
                                        if (mabl.ContainsKey(bo))
                                        {
                                            foreach (var item3 in codes.Keys)
                                            {
                                                if (mabl[bo].ContainsKey(item3))
                                                {
                                                    mabl[bo][item3] = mabl[bo][item3] + codes[item3];
                                                }
                                                else
                                                {
                                                    mabl[bo].Add(item3, codes[item3]);
                                                }
                                            }
                                        }
                                        else
                                        {

                                            mabl.Add(bo, codes);

                                        }
                                    }

                                }

                                tr.Commit();
                            }

                            if (mabl.Count != 0)
                            {
                                OpenFileDialog openDialog = new OpenFileDialog();
                                openDialog.Multiselect = true;
                                if (DialogResult.OK == openDialog.ShowDialog())
                                {
                                    string[] filename = openDialog.FileNames;

                                    // EXLf eXLf = new EXLf();
                                    //   res = eXLf.evW(mabl, filename);
                                    NPOIExcel nPOIExcel = new NPOIExcel(Path.GetDirectoryName(doc.Name));
                                    res = nPOIExcel.Calv(mabl, filename);
                                }

                            }

                        }

                        if (res)
                        {
                            megbox.Show("数量一致！");
                        }
                        else
                        { megbox.Show("数量不一致！"); }


                    }
                }
            }

            catch (System.Exception E)
            {

                MyPlugin.ExceptionWrit(E);
            }
        }

        /// <summary>
        /// 自动添加锁条
        /// </summary>
        [CommandMethod("ADDS")]
        public void MyCommand7() // This method can have any name
        {
            try
            {
                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                //   Point3d loca=new Point3d ( );
                DBObjectCollection dc = new DBObjectCollection();

                ObjectIdCollection oj = new ObjectIdCollection();

                Dictionary<string, BlockReference> mabl = new Dictionary<string, BlockReference>();

                using (doc.LockDocument())
                {
                    //  TypedValue[] tl=new TypedValue[]{new TypedValue((int)DxfCode.Operator,"<AND"),new TypedValue((int)DxfCode.Start,"DBText"),new TypedValue((int)DxfCode.Operator,"AND>")};
                    TypedValue[] tl = new TypedValue[] { new TypedValue(0, "Text") };
                    SelectionFilter fu = new SelectionFilter(tl);
                    PromptSelectionOptions selop = new PromptSelectionOptions();
                    PromptSelectionResult sel = ed.SelectAll(fu);

                    if (sel.Status == PromptStatus.OK)
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            SelectionSet seva = sel.Value;
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            foreach (var item in seva.GetObjectIds())
                            {
                                DBText dtex = (DBText)tr.GetObject(item, OpenMode.ForRead);

                                if (dtex.TextString == "Z(3)")
                                {
                                    DBText t1 = (DBText)dtex.Clone();
                                    DBText t2 = (DBText)dtex.Clone();

                                    if ((t1.Rotation >= -0.1) && (t1.Rotation <= 0.1))
                                    {
                                        t1.AlignmentPoint = new Point3d(dtex.AlignmentPoint.X, dtex.AlignmentPoint.Y + dtex.Height + (0.1 * dtex.Height), dtex.AlignmentPoint.Z);
                                        t2.AlignmentPoint = new Point3d(dtex.AlignmentPoint.X, dtex.AlignmentPoint.Y - dtex.Height - (0.1 * dtex.Height), dtex.AlignmentPoint.Z);
                                    }
                                    else
                                    {
                                        t1.AlignmentPoint = new Point3d(dtex.AlignmentPoint.X + dtex.Height + (0.1 * dtex.Height), dtex.AlignmentPoint.Y, dtex.AlignmentPoint.Z);
                                        t2.AlignmentPoint = new Point3d(dtex.AlignmentPoint.X - dtex.Height - (0.1 * dtex.Height), dtex.AlignmentPoint.Y, dtex.AlignmentPoint.Z);
                                    }

                                    t1.TextString = "ST";
                                    t2.TextString = "ST";

                                    br.AppendEntity(t1);
                                    tr.AddNewlyCreatedDBObject(t1, true);
                                    br.AppendEntity(t2);
                                    tr.AddNewlyCreatedDBObject(t2, true);
                                }

                            }

                            tr.Commit();
                        }
                    }

                }
            }
            catch (System.Exception E)
            {
                MyPlugin.ExceptionWrit(E);
            }

        }

        [CommandMethod("getAr")]
        public void MyCommand9() // This method can have any name
        {
            try
            {
                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                double are = 0;
                double le = 0;

                using (doc.LockDocument())
                {
                    //  TypedValue[] tl=new TypedValue[]{new TypedValue((int)DxfCode.Operator,"<AND"),new TypedValue((int)DxfCode.Start,"DBText"),new TypedValue((int)DxfCode.Operator,"AND>")};
                    TypedValue[] tl = new TypedValue[] { new TypedValue(0, "LWPolyLine") };
                    SelectionFilter fu = new SelectionFilter(tl);
                    PromptSelectionOptions selop = new PromptSelectionOptions();
                    PromptSelectionResult sel = ed.GetSelection(selop, fu);

                    if (sel.Status == PromptStatus.OK)
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            SelectionSet seva = sel.Value;
                            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                            foreach (var item in seva.GetObjectIds())
                            {
                                Polyline dtex = (Polyline)tr.GetObject(item, OpenMode.ForRead);

                                if (dtex.Closed)
                                {
                                    are += dtex.Area;
                                }
                                le += dtex.Length;
                            }

                            tr.Commit();
                        }
                    }

                }

                ed.WriteMessage("总面积是：" + are.ToString() + "；" + "总长度是：" + le.ToString() + "；\n");
            }
            catch (System.Exception E)
            {
                MyPlugin.ExceptionWrit(E);
            }

        }

        [CommandMethod("Cdwg")]
        public void MyCommand20() // This method can have any name
        {
            try
            {
                Document doc = AApplication.DocumentManager.MdiActiveDocument;

                Editor ed = doc.Editor;

                Database db = doc.Database;

                double H = 1, T, R;

                List<Point3d> point3Ds = new List<Point3d>();

                using (doc.LockDocument())
                {
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                        foreach (var item in br)
                        {

                            var dd = tr.GetObject(item, OpenMode.ForRead);
                            if (dd is DBText)
                            {
                                DBText dtex = (DBText)dd;

                                point3Ds.Add(new Point3d((dtex.GeometricExtents.MaxPoint.X + dtex.GeometricExtents.MinPoint.X) / 2,
                                    (dtex.GeometricExtents.MaxPoint.Y + dtex.GeometricExtents.MinPoint.Y) / 2, 0));

                                if (Regex.IsMatch(dtex.TextString, "^(JD|Q|M|L|DM|T)-\\d{1,3}$"))
                                {
                                    H = dtex.Height;
                                }

                            }

                        }

                        Line lineL = new Line(point3Ds[0], point3Ds[1]);

                        T = lineL.Length / H;
                        R = lineL.Angle * 180 / Math.PI;
                        tr.Commit();
                    }
                }

            }
            catch (System.Exception E)
            {
                MyPlugin.ExceptionWrit(E);
            }

        }

        [CommandMethod("TuC")]
        public void MyCommand12()
        {
            Document doc = AApplication.DocumentManager.MdiActiveDocument;

            Editor ed = doc.Editor;

            Database db = doc.Database;

            using (doc.LockDocument())
            {
                LayerTableRecord layeyscordm = new LayerTableRecord();
                LayerTableRecord layeyscordm1 = new LayerTableRecord();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable layeys = tr.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;
                    LinetypeTable linetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForWrite) as LinetypeTable;

                    foreach (var item in layeys)
                    {
                        LayerTableRecord layeyscord = tr.GetObject(item, OpenMode.ForWrite) as LayerTableRecord;
                        if (Regex.IsMatch(layeyscord.Name, "轮廓"))
                        {
                            layeyscord.LineWeight = LineWeight.LineWeight030;

                        }
                    }

                    if (!layeys.Has("虚线"))
                    {
                        LayerTableRecord layeyscord = new LayerTableRecord();
                        layeyscord.LineWeight = LineWeight.LineWeight015;
                        layeyscord.Name = "虚线";
                        layeyscord.Color = Autodesk.AutoCAD.Colors.Color.FromColor(System.Windows.Media.Colors.Yellow);

                        if (linetypeTable.Has("Hidden"))
                        {
                            layeyscord.LinetypeObjectId = linetypeTable["Hidden"];

                        }

                        layeys.Add(layeyscord);
                        tr.AddNewlyCreatedDBObject(layeyscord, true);
                    }

                    tr.Commit();
                }
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable layeys = tr.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;

                    foreach (var item in layeys)
                    {
                        LayerTableRecord layeyscord = tr.GetObject(item, OpenMode.ForWrite) as LayerTableRecord;
                        if (Regex.IsMatch(layeyscord.Name, "轮廓"))
                        {
                            layeyscordm1 = layeyscord;

                        }
                        else if (Regex.IsMatch(layeyscord.Name, "虚线"))
                        {
                            layeyscordm = layeyscord;
                        }
                    }

                    tr.Commit();
                }

                TypedValue[] t2 = new TypedValue[] { new TypedValue((int)DxfCode.Operator, "<or"),
                    new TypedValue(0, "POLYLINE"),new TypedValue(0, "LWPOLYLINE"),new TypedValue(0, "Circle"),
                    new TypedValue(0, "Line"),new TypedValue(0, "Arc"),new TypedValue(0, "Hatch"),
                    new TypedValue(0, "Curve"),new TypedValue(0, "Ellipse"), new TypedValue(0, "TEXT"),
                new TypedValue((int)DxfCode.Operator, "or>")};

                SelectionFilter fu = new SelectionFilter(t2);
                PromptSelectionOptions selop = new PromptSelectionOptions();
                PromptSelectionResult sel = ed.GetSelection(selop, fu);

                if (sel.Status == PromptStatus.OK)
                {
                    List<ObjectId> ids = new List<ObjectId>();
                    List<int> wp = new List<int>();

                    List<ObjectId> idsha = new List<ObjectId>();

                    Dictionary<Point3d, string> dhapo = new Dictionary<Point3d, string>();

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {

                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                        BlockTableRecord br = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        LinetypeTable linetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForWrite) as LinetypeTable;

                        SelectionSet seva = sel.Value;

                        foreach (var item in seva.GetObjectIds())
                        {
                            Entity dtexM = tr.GetObject(item, OpenMode.ForWrite) as Entity;

                            if (dtexM is Hatch)
                            {
                                Hatch hat = dtexM as Hatch;
                                if (hat.PatternName != "SOLID")
                                {
                                    idsha.Add(item);
                                    Point3d mah = hat.GeometricExtents.MaxPoint;
                                    Point3d mih = hat.GeometricExtents.MinPoint;
                                    Point3d ph = new Point3d((mah.X + mih.X) / 2, (mah.Y + mih.Y) / 2, 0);

                                    while (dhapo.ContainsKey(ph))
                                    {
                                        ph = new Point3d(ph.X - 0.01, ph.Y - 0.01, 0);
                                    }
                                    dhapo.Add(ph, hat.PatternName);
                                }
                            }
                            else if (dtexM is Polyline3d && dtexM.Layer == "轮廓")
                            {
                                Polyline3d dtex = (Polyline3d)tr.GetObject(item, OpenMode.ForWrite);

                                DBObjectCollection dBObject1 = new DBObjectCollection();

                                dtex.Explode(dBObject1);

                                if ((dBObject1.Count == 4 || dBObject1.Count == 3) && Isclose(dtex.StartPoint, dtex.EndPoint))
                                {
                                    ids.Add(item);
                                }
                                else
                                {
                                    dtexM.SetLayerId(layeyscordm1.ObjectId, true);
                                    dtexM.LineWeight = layeyscordm1.LineWeight;
                                    dtexM.LinetypeId = layeyscordm1.LinetypeObjectId;
                                    dtexM.Color = layeyscordm1.Color;
                                }
                            }
                            else if (dtexM.Linetype == "Dashdot")
                            {
                                dtexM.SetLayerId(layeyscordm.ObjectId, true);
                                dtexM.LinetypeId = layeyscordm.LinetypeObjectId;
                                dtexM.LineWeight = layeyscordm.LineWeight;
                                dtexM.Color = layeyscordm.Color;
                                dtexM.LinetypeScale = 0.5;
                            }
                            else if (dtexM is Polyline && dtexM.Layer == "轮廓")
                            {
                                dtexM.SetLayerId(layeyscordm1.ObjectId, true);
                                dtexM.LineWeight = layeyscordm1.LineWeight;
                                dtexM.LinetypeId = layeyscordm1.LinetypeObjectId;
                                dtexM.Color = layeyscordm1.Color;
                            }
                            else if (dtexM is Arc && dtexM.Layer == "孔位")
                            {
                                dtexM.SetLayerId(layeyscordm1.ObjectId, true);
                                dtexM.LineWeight = layeyscordm1.LineWeight;
                                dtexM.LinetypeId = layeyscordm1.LinetypeObjectId;
                                dtexM.Color = layeyscordm1.Color;
                            }
                            else if (dtexM is Ellipse || dtexM is Circle)
                            {
                                dtexM.SetLayerId(layeyscordm1.ObjectId, true);
                                dtexM.LineWeight = layeyscordm1.LineWeight;
                                dtexM.LinetypeId = layeyscordm1.LinetypeObjectId;
                                dtexM.Color = layeyscordm1.Color;
                            }
                            else if (dtexM is DBText && dtexM.Layer == "标注")
                            {
                                DBText bText = dtexM as DBText;
                                string dwm = bText.TextString.Replace(" ", "");
                                if (Regex.IsMatch(dwm, "^(贴片|企口|滴水)"))
                                {
                                    string dw = Regex.Match(dwm, "\\d{1,2}×\\d{2,3}").Value;
                                    if (Regex.IsMatch(dw, "\\d{2,3}$"))
                                    {
                                        int dwi = Convert.ToInt32(Regex.Match(dw, "\\d{2,3}$").Value);

                                        int laz = dwi % 10;
                                        switch (laz)
                                        {
                                            case 9:
                                                dwi = dwi + 1;
                                                break;
                                            case 8:
                                                dwi = dwi + 2;
                                                break;
                                            case 7:
                                                dwi = dwi + 3;
                                                break;
                                            default:
                                                break;
                                        }
                                        if (!wp.Contains(dwi)) wp.Add(dwi);
                                    }
                                }
                            }
                        }

                        foreach (ObjectId item in ids)
                        {
                            DBObjectCollection dBObject1 = new DBObjectCollection();
                            Polyline3d dtexM = (Polyline3d)tr.GetObject(item, OpenMode.ForWrite);
                            dtexM.Explode(dBObject1);
                            int va = 0;
                            int va1 = 0;
                            foreach (var item1 in dBObject1)
                            {
                                if (item1 is Line)
                                {
                                    Line line = item1 as Line;
                                    int laz1 = (int)Math.Round(line.Length, 0);
                                    int laz = laz1 % 10;
                                    if (laz == 9)
                                    {
                                        laz = laz1 + 1;
                                    }
                                    else if (laz == 1)
                                    {
                                        laz = laz1 - 1;
                                    }
                                    else
                                    {
                                        laz = laz1;
                                    }
                                    if (laz == 20)
                                    {
                                        va1 += 1;
                                        break;
                                    }
                                }
                            }

                            if (va1 == 1)
                            {
                                List<string> hap = new List<string>();
                                foreach (Point3d hp in dhapo.Keys)
                                {
                                    if (PositionPnpoly(dBObject1, hp.X, hp.Y))
                                    {
                                        hap.Add(dhapo[hp]);
                                    }
                                }
                                if (hap.Contains("DOLMIT"))
                                {
                                    va += 1;
                                }
                                else if (hap.Contains("AR-CONC"))
                                {
                                    va = 0;
                                }
                            }
                            else
                            {
                                foreach (var item1 in dBObject1)
                                {
                                    if (item1 is Line)
                                    {
                                        Line line = item1 as Line;
                                        int laz1 = (int)Math.Round(line.Length, 0);
                                        int laz = laz1 % 10;

                                        if (laz == 9)
                                        {
                                            laz = laz1 + 1;
                                        }
                                        else if (laz == 1)
                                        {
                                            laz = laz1 - 1;
                                        }
                                        else
                                        {
                                            laz = laz1;
                                        }
                                        if (laz <= 30)
                                        {
                                            va = 0;
                                            dtexM.SetLayerId(layeyscordm1.ObjectId, true);
                                            dtexM.LineWeight = layeyscordm1.LineWeight;
                                            dtexM.LinetypeId = layeyscordm1.LinetypeObjectId;
                                            dtexM.Color = layeyscordm1.Color;
                                            break;
                                        }
                                        else if (laz > 30 && wp.Contains(laz))
                                        {
                                            va += 1;

                                        }

                                    }
                                }
                            }

                            if (va == 0)
                            {
                                dtexM.SetLayerId(layeyscordm1.ObjectId, true);
                                dtexM.LineWeight = layeyscordm1.LineWeight;
                                dtexM.LinetypeId = layeyscordm1.LinetypeObjectId;
                                dtexM.Color = layeyscordm1.Color;
                            }
                            else
                            {
                                dtexM.SetLayerId(layeyscordm.ObjectId, true);
                                dtexM.LinetypeId = layeyscordm.LinetypeObjectId;
                                dtexM.LineWeight = layeyscordm.LineWeight;
                                dtexM.Color = layeyscordm.Color;
                                dtexM.LinetypeScale = 0.5;
                                DBObjectCollection dBObject2 = new DBObjectCollection();
                                dtexM.Explode(dBObject2);
                                dtexM.Erase(true);
                                foreach (var item1 in dBObject2)
                                {
                                    if (item1 is Line)
                                    {
                                        Line line = item1 as Line;
                                        br.AppendEntity(line);
                                        tr.AddNewlyCreatedDBObject(line, true);
                                    }
                                }
                            }

                            foreach (ObjectId item3 in idsha)
                            {
                                Hatch dtex3 = (Hatch)tr.GetObject(item3, OpenMode.ForWrite);
                                dtex3.Erase(true);
                            }

                        }
                        tr.Commit();
                    }
                }
            }
        }
        private Point3dCollection getDBp(DBText dBText, double r)
        {
            Point3dCollection point3DCollection;

            double rot = dBText.Rotation;

            Point3d pmin, pmax1, pminax1, pmaxin1;

            if (rot != 0)
            {
                DBText dBText2 = (DBText)dBText.Clone();
                dBText2.Rotation = 0;

                double x = dBText2.GeometricExtents.MaxPoint.X - dBText2.GeometricExtents.MinPoint.X;
                double y = dBText2.GeometricExtents.MaxPoint.Y - dBText2.GeometricExtents.MinPoint.Y;

                pmin = new Point3d(dBText2.GeometricExtents.MinPoint.X + x * r,
                    dBText2.GeometricExtents.MinPoint.Y + y * r, 0);

                Point3d pmax = new Point3d(dBText2.GeometricExtents.MaxPoint.X - dBText2.GeometricExtents.MinPoint.X - x * r,
                   dBText2.GeometricExtents.MaxPoint.Y - dBText2.GeometricExtents.MinPoint.Y - y * r, 0);

                Point3d pminax = new Point3d(0,
                    dBText2.GeometricExtents.MaxPoint.Y - dBText2.GeometricExtents.MinPoint.Y - y * r, 0);

                Point3d pmaxin = new Point3d(dBText2.GeometricExtents.MaxPoint.X - dBText2.GeometricExtents.MinPoint.X - x * r,
                   0, 0);

                pmax1 = new Point3d(pmax.X * Math.Cos(rot) - pmax.Y * Math.Sin(rot) + pmin.X,
               pmax.X * Math.Sin(rot) + pmax.Y * Math.Cos(rot) + pmin.Y, 0);

                pminax1 = new Point3d(pminax.X * Math.Cos(rot) - pminax.Y * Math.Sin(rot) + pmin.X,
             pminax.X * Math.Sin(rot) + pminax.Y * Math.Cos(rot) + pmin.Y, 0);

                pmaxin1 = new Point3d(pmaxin.X * Math.Cos(rot) - pmaxin.Y * Math.Sin(rot) + pmin.X,
           pmaxin.X * Math.Sin(rot) + pmaxin.Y * Math.Cos(rot) + pmin.Y, 0);

            }
            else
            {
                double x = dBText.GeometricExtents.MaxPoint.X - dBText.GeometricExtents.MinPoint.X;
                double y = dBText.GeometricExtents.MaxPoint.Y - dBText.GeometricExtents.MinPoint.Y;

                pmin = new Point3d(dBText.GeometricExtents.MinPoint.X + x * r,
                  dBText.GeometricExtents.MinPoint.Y + y * r, 0);

                pmax1 = new Point3d(dBText.GeometricExtents.MaxPoint.X - x * r,
                  dBText.GeometricExtents.MaxPoint.Y - y * r, 0);

                pminax1 = new Point3d(pmin.X, pmax1.Y, 0);

                pmaxin1 = new Point3d(pmax1.X, pmin.Y, 0);
            }

            point3DCollection = new Point3dCollection();

            point3DCollection.Add(pmin);
            point3DCollection.Add(pminax1);
            point3DCollection.Add(pmax1);
            point3DCollection.Add(pmaxin1);

            return point3DCollection;
        }

        private Dictionary<string, int> getDBm(string code1, Dictionary<string, string> mapcode, string[] strcode1s, string pat)
        {
            Dictionary<string, int> PD = new Dictionary<string, int>();

            if (code1 != "")
            {
                string suf = "";

                code1 = code1.Replace(" ", "");
                code1 = Regex.Replace(code1, "\\(?[\u4e00-\u9fa5]+\\)?", "");

                if (Regex.IsMatch(code1, "^([MLQDTJ]|[MLQDTJ]B\\d)-"))
                {
                    suf = Regex.Match(code1, "^([MLQDTJ]|[MLQDTJ]B\\d)-").Value;
                    code1 = Regex.Replace(code1, "^([MLQDTJ]|[MLQDTJ]B\\d)-", "");
                }

                if (code1 == "") return PD;

                string lcoded = "";

                for (int i = 0; i < strcode1s.Length; i++)
                {
                    lcoded = Regex.Match(code1, strcode1s[i]).Value;

                    if (lcoded != "") break;

                }

                lcoded = Regex.Replace(lcoded, "\\d{1,3}\\.?\\d{0,3}", "#");

                if (lcoded != "" && mapcode.ContainsKey(lcoded))
                {

                    if (Regex.IsMatch(code1, mapcode[lcoded]))
                    {
                        string code3 = code1;
                        int nub = 1;
                        if (Regex.IsMatch(code1, pat))
                        {
                            string codenub = Regex.Match(code1, pat).Value;
                            nub = Convert.ToInt32(Regex.Match(codenub, "\\d").Value);
                            code3 = Regex.Replace(code1, pat, "");
                        }
                        else if (Regex.IsMatch(code1, "(\\([23456]\\)|\\[[23456]\\])$"))
                        {
                            nub = 1;
                            code3 = Regex.Replace(code1, "(\\([23456]\\)|\\[[23456]\\])$", "");
                        }

                        if (suf != "") code3 = suf + code3;

                        PD.Add(code3, nub);

                    }

                }

            }

            return PD;

        }

        /// <summary>
        /// 判断当前位置是否在不规则形状里面
        /// </summary>
        /// <param name="nvert">不规则形状的定点数</param>
        /// <param name="vertx">不规则形状x坐标集合</param>
        /// <param name="verty">不规则形状y坐标集合</param>
        /// <param name="testx">当前x坐标</param>
        /// <param name="testy">当前y坐标</param>
        /// <returns></returns>
        public static bool PositionPnpoly(int nvert, List<double> vertx, List<double> verty, double testx, double testy)
        {
            int i, j, c = 0;
            for (i = 0, j = nvert - 1; i < nvert; j = i++)
            {
                if (((verty[i] > testy) != (verty[j] > testy)) && (testx < (vertx[j] - vertx[i]) * (testy - verty[i]) / (verty[j] - verty[i]) + vertx[i]))
                {
                    c = 1 + c; ;
                }
            }
            if (c % 2 == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        /// <summary>
        /// 判断当前位置是否在不规则形状里面
        /// </summary>
        /// <param name="lines">不规则形状直线集合</param>
        /// <param name="testx">当前x坐标</param>
        /// <param name="testy">当前y坐标</param>
        /// <returns></returns>
        public static bool PositionPnpoly(List<Line> lines, double testx, double testy)
        {
            int c = 0;

            foreach (Line item in lines)
            {
                Point3d starp = item.StartPoint;
                Point3d endp = item.EndPoint;

                if (((starp.Y > testy) != (endp.Y > testy)) && (testx < (starp.X - endp.X) * (testy - endp.Y) / (starp.Y - endp.Y) + endp.X))
                {
                    c = 1 + c; ;
                }
            }

            if (c % 2 == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// 判断当前位置是否在不规则形状里面
        /// </summary>
        /// <param name="lines">不规则形状直线集合</param>
        /// <param name="testx">当前x坐标</param>
        /// <param name="testy">当前y坐标</param>
        /// <returns></returns>
        public static bool PositionPnpoly(DBObjectCollection lines, double testx, double testy)
        {
            int c = 0;

            foreach (Line item in lines)
            {
                Point3d starp = item.StartPoint;
                Point3d endp = item.EndPoint;

                if (((starp.Y > testy) != (endp.Y > testy)) && (testx < (starp.X - endp.X) * (testy - endp.Y) / (starp.Y - endp.Y) + endp.X))
                {
                    c = 1 + c; ;
                }
            }

            if (c % 2 == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public static Point3dCollection gePoly(Point3d point3D)
        {
            Point3dCollection pc = new Point3dCollection();

            pc.Add(new Point3d(point3D.X - 0.6, point3D.Y - 0.6, 0));
            pc.Add(new Point3d(point3D.X - 0.6, point3D.Y + 0.6, 0));
            pc.Add(new Point3d(point3D.X + 0.6, point3D.Y + 0.6, 0));
            pc.Add(new Point3d(point3D.X + 0.6, point3D.Y - 0.6, 0));

            return pc;

        }

        public static bool Isclose(Point3d stpoint3D, Point3d edpoint3D)
        {

            if (Math.Round(stpoint3D.X - edpoint3D.X, 3) == 0 && Math.Round(stpoint3D.Y - edpoint3D.Y, 3) == 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }
    }
}
