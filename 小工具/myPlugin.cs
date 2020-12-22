// (C) Copyright 2018 by Microsoft 
//
using System;
using Autodesk . AutoCAD . Runtime;
using Autodesk . AutoCAD . ApplicationServices;
using Autodesk . AutoCAD . DatabaseServices;
using Autodesk . AutoCAD . Geometry;
using Autodesk . AutoCAD . EditorInput;
//using Autodesk . AutoCAD . Interop;
using Autodesk . AutoCAD . Windows;
using System . IO;
using System . Reflection;
using System . Data . OleDb;
using Exceptions=System.Exception;
// This line is not mandatory, but improves loading performances
[assembly: ExtensionApplication ( typeof ( 小工具 . MyPlugin ) )]

namespace 小工具
{

    // This class is instantiated by AutoCAD once and kept alive for the 
    // duration of the session. If you don't do any one time initialization 
    // then you should remove this class.
    public class MyPlugin : IExtensionApplication
    {
        public static string direct;
        MyCommands comm=new MyCommands ( );
        void IExtensionApplication . Initialize ( )
        {
            direct = Path . GetDirectoryName ( Assembly . GetExecutingAssembly ( ) . Location );
            ContextMenuExtension come=new ContextMenuExtension ( );
            
            come . Title = "小工具";
           // MenuItem mi=new MenuItem ( "自动图纸编号" );
          //  mi . Click += mi_Click; ;
          
          //  MenuItem mi1=new MenuItem ( "编号对照表" );
          //  mi1 . Click += mi1_Click;
      
         //   MenuItem mi2=new MenuItem ( "清除图框编号" );
         //   mi2 . Click += mi2_Click;
           
            MenuItem mi3=new MenuItem ("核对分区数量(CFQ)");
            mi3.Click+=mi3_Click;
          
           // MenuItem mi4=new MenuItem ( "生产图纸" );
           // mi4 . Click += mi4_Click;
         
            MenuItem mi5=new MenuItem ("核对数量(CNU)");
            mi5 . Click += mi5_Click;

            MenuItem mi7 = new MenuItem("检查文字重叠(CDB)");
            mi7.Click += Mi7_Click;

          //  MenuItem mi8 = new MenuItem("获取面积和周长");
          //  mi8.Click += Mi8_Click; 

            MenuItem mi6 =new MenuItem ("添加ST(ADDS)");
            mi6 . Click += mi6_Click;

            MenuItem mi10= new MenuItem("处理图纸(TUC)");
            mi10.Click += Mi10_Click;
            //  come . MenuItems . Add ( mi );
            //  come . MenuItems . Add ( mi1 );
            //  come . MenuItems . Add ( mi2 );
            come . MenuItems . Add ( mi3 );
           // come . MenuItems . Add ( mi4 );
            come . MenuItems . Add ( mi5 );
            come . MenuItems . Add ( mi6 );
         //   come.MenuItems.Add(mi8);
            come.MenuItems.Add(mi7);

            come.MenuItems.Add(mi10);

            Autodesk . AutoCAD . ApplicationServices . Application . AddDefaultContextMenuExtension ( come );
        }

        private void Mi10_Click(object sender, EventArgs e)
        {
            comm.MyCommand12();
        }

        //  private void Mi8_Click(object sender, EventArgs e)
        //  {
        //      comm.MyCommand9();
        //  }

        private void Mi7_Click(object sender, EventArgs e)
        {
            comm.MyCommand8();
        }

        void mi6_Click ( object sender , EventArgs e )
        {
            comm . MyCommand7( );
        }

        void mi5_Click ( object sender , EventArgs e )
        {
            comm.MyCommand6();
        }

      //  void mi4_Click ( object sender , EventArgs e )
      //  {
      //     comm.MyCommand5();
     //   }
        void mi3_Click ( object sender , EventArgs e )
        {
            comm .MyCommand10();
        }

        //void mi2_Click ( object sender , EventArgs e )
      //  {
      //    comm.MyCommand4();
      //  }

      //  void mi1_Click ( object sender , EventArgs e )
      //  {
     //      comm.MyCommand1();
     //   }

      //  void mi_Click ( object sender , EventArgs e )
      //  {
      //      comm.MyCommand();
     //   }

        public static void ExceptionWrit(Exceptions ex)
        {
            string path = "D:\\CadExceptionLog.txt";
            FileStream fs;
            StreamWriter sw;
            long fl;

            if (File.Exists(path))
            {
                fs = new FileStream(path, FileMode.Open, FileAccess.Write);
                sw = new StreamWriter(fs,System.Text.UTF8Encoding.Unicode);
                fl = fs.Length;
                fs.Seek(fl, SeekOrigin.Begin);
            }
            else
            {
                fs = new FileStream(path, FileMode.Create, FileAccess.Write);
                sw = new StreamWriter(fs, System.Text.UTF8Encoding.Unicode);
                fl = fs.Length;
                fs.Seek(fl, SeekOrigin.End);
            }

            sw.WriteLine(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + ";");
            sw.WriteLine(ex.Message + ex.StackTrace);
            sw.WriteLine();

            sw.Close();
            fs.Close();

        }

        void IExtensionApplication . Terminate ( )
        {
            // Do plug-in application clean up here
        }
        /// <summary>
        /// 注册表注册插件，完成自动加载
        /// </summary>
      
    }

}
