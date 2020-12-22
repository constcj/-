using System;
using System . Collections . Generic;
using System . Linq;
using System . Text;
using System . Threading . Tasks;
using Microsoft . Office . Interop . Excel;
using System . Diagnostics;
using System . IO;
using System . Text . RegularExpressions;
using ExlApp = Microsoft . Office . Interop . Excel . Application;
using System . Windows . Forms;

using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;

namespace 小工具
{
    class EXL
    {
        private  Workbook wbook=null;
        private  Worksheet wsheet=null;
        private    ExlApp exl;
         private  int h;
        private static int ope=0;
     public   EXL()
        {
            try
            {
   
            int l = Process . GetProcessesByName ( "EXCEL" ) . Length;
            if ( l > 0 )
            {
                exl = ( ExlApp ) System . Runtime . InteropServices . Marshal . GetActiveObject ( "Excel.Application" );
                exl . Visible = false;
            }
            else
            {
                exl= new Microsoft . Office . Interop . Excel . Application ( );
                exl . Visible = false;
            }
            string pa="F:\\图纸编号";
            if ( !Directory . Exists (pa) )
            {
                Directory . CreateDirectory ( pa );
                wbook = exl . Workbooks . Add ( );
                wbook . SaveAs ( pa + @"\图纸编号" , XlFileFormat . xlOpenXMLWorkbook );
                wsheet = wbook . Sheets [ "Sheet1" ] as Worksheet;
            }
            else if ( !File . Exists ( pa + @"\图纸编号.xlsx" ) )
            {
                wbook = exl . Workbooks . Add ( );
                wbook . SaveAs ( pa + @"\图纸编号" , XlFileFormat . xlWorkbookNormal );
                wsheet = wbook . Sheets [ "Sheet1" ] as Worksheet;
            }
            else
            {
                foreach ( Workbook item in exl.Workbooks )
                {
                    if ( item.Name=="图纸编号" )
                    {
                          wbook = exl . Workbooks [ "图纸编号" ];
                          ope = 1;
                    }
                }
                if ( wbook==null )
                {
                    wbook = exl . Workbooks . Open ( pa + @"\图纸编号.xlsx" );
                }            
                wsheet = wbook . Sheets [ "Sheet1" ] as Worksheet;
               
            }
          wsheet . get_Range ( "a1" ) . Value ="图号";
          wsheet . get_Range ( "b1" ) . Value = "模板编号";

            }
            catch ( Exception E)
            {

                MyPlugin.ExceptionWrit(E);

            }   
        }
          private void GetEndRow ( string c)
        {
            h = wsheet . get_Range ( c+65536 ) . get_End ( XlDirection . xlUp ) . Row + 1;}
        public void SetV(string code,string tuhao)
        {
            GetEndRow ( "a" );
            wsheet . get_Range ( "a" + h ) . Value = tuhao;
            wsheet . get_Range ( "b" + h ) . Value =code;
        }
        public void Col()
        {
            wbook . Save ( );
            if ( ope==0 )
            {
                wbook . Close ( );
            } 
        }
    }
    class EXLe
    {
        private  Workbook wbook=null;
        private  Worksheet wsheet=null;
        private    ExlApp exl;
        private  int h;
        private static int ope=0;
        public EXLe (string na )
        {
            try
            {

            int l = Process . GetProcessesByName ( "EXCEL" ) . Length;
            if ( l > 0 )
            {
                exl = ( ExlApp ) System . Runtime . InteropServices . Marshal . GetActiveObject ( "Excel.Application" );
                exl . Visible = false;
            }
            else
            {
                exl = new Microsoft . Office . Interop . Excel . Application ( );
                exl . Visible = false;
            }     
                    wbook = exl . Workbooks . Open (na);
              
            }
            catch ( Exception E)
            {
                MyPlugin.ExceptionWrit(E);
            }     
        }
        private void GetEndRow ( string c )
        {
            h = wsheet . get_Range ( c + 65536 ) . get_End ( XlDirection . xlUp ) . Row;
        }
      public SortedSet<string> GetDar()
        {
            SortedSet<string> sed=new SortedSet<string> ( );
            try
            {

            string cl="J";
            int rh=7;
            foreach ( Worksheet item in wbook.Worksheets )
            {
                wsheet = item;
                object v=wsheet . get_Range ( "j7") . Value;
                if (v!=null && v.ToString()=="图纸编号" )
                {
                    rh = 9;
                }else
                { rh = 7;
                cl = "i";
                }
                GetEndRow ( "b" );

                for ( int i = rh ; i < h+1 ; i++ )
                {
                    object dar=wsheet . get_Range ( cl + i ) . Value;
                
                   if ( dar!=null)
                   {
                       string th=dar . ToString ( );

                       if (!Regex.IsMatch(th,pattern:"^ZW-[CQPNZKJTF]-" ))
                       {
                           sed . Add ( dar . ToString ( ) );
                       }                    
                   } 
                }

            }
            }
            catch ( Exception E)
            {
                MyPlugin.ExceptionWrit(E);
            }
            return sed;
        }
      public SortedSet<string> GetDar1 ( )
      {
          SortedSet<string> sed=new SortedSet<string> ( );
          try
          {

          int rh=7;
          foreach ( Worksheet item in wbook . Worksheets )
          {
              wsheet = item;
              object v=wsheet . get_Range ( "j7" ) . Value;
              if ( v != null && v . ToString ( ) == "图纸编号" )
              {
                  rh = 9;
              }
              else
              {
                  rh = 7;
              }
              GetEndRow ( "b" );

              for ( int i = rh ; i < h + 1 ; i++ )
              {
                  object dar=wsheet . get_Range ( "c" + i ) . Value;

                  if ( dar != null )
                  {
                      string th=dar . ToString ( );
                       sed . Add ( dar . ToString ( ) );
                  }
              }

          }

          }
          catch ( Exception E)
          {

                MyPlugin.ExceptionWrit(E);
          }
          return sed;
      }
        public Dictionary<string,string> GetT()
      {
          wsheet = wbook . Sheets [ "Sheet1" ] as Worksheet;
            Dictionary<string,string> dii=new Dictionary<string,string>();
          GetEndRow ( "b" );
          if ( h>2 )
          {     
          for ( int i = 2; i < h + 1 ; i++ )
          {
              object dar=wsheet . get_Range ( "b" + i ) . Value;
              object tu=wsheet . get_Range ( "a" + i ) . Value;
              if ( dar != null && tu!=null)
              {
                  string th=Regex.Replace( dar . ToString ( ),pattern:"\\([A-Z0-9\\.]+\\)",replacement:"");

                  if ( !dii.ContainsKey(th) )
                  {
                      dii . Add ( th , tu . ToString ( ) );
                  } else
                  { MessageBox.Show(dar.ToString()+"为模板重复编码");}
              }
          }
               
          }
            return dii;
      }
      public void Col ( )
      {
          wbook . Save ( );
            wbook . Close ( );

      }
    }
    class EXLg
    {
        private Workbook wbook = null;
        private Worksheet wsheet = null;
        private ExlApp exl;
        private int h;
        public EXLg(ref Boolean hav)
        {
           
            try
            {

                int l = Process.GetProcessesByName("EXCEL").Length;
                if (l > 0)
                {
                    ExlApp exl1 = (ExlApp)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    foreach (Workbook item in exl1.Workbooks)
                    {
                        if (Regex.IsMatch(item.Name,"^总量\\.xls"))
                        {
                            Workbook wbook1 = item;
                            foreach (Worksheet item1 in wbook1.Sheets)
                            {
                                if (item1.Name=="核对数量")
                                {
                                    hav = true;
                                }
                            }
                        }
                    }

                }
              
            }
            catch (Exception E)
            {

                MyPlugin.ExceptionWrit(E);

            }
 
        }

        public void setValueRange(Dictionary<string, int> dic)
        {

            try
            {

                int l = Process.GetProcessesByName("EXCEL").Length;
                if (l > 0)
                {
                    exl = (ExlApp)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    foreach (Workbook item2 in exl.Workbooks)
                    {
                        for (int k =1; k < exl.Workbooks.Count+1; k++)
                        {
                            if (Regex.IsMatch(exl.Workbooks.Item[k].Name, "^总量\\.xls"))
                            {
                                wbook = exl.Workbooks.Item[k];

                                for (int g =1; g < wbook.Sheets.Count+1; g++)
                                {
                                    if (wbook.Sheets.Item[g].Name == "核对数量")
                                    {
                                        wsheet = wbook.Sheets.Item[g];

                                        h = wsheet.get_Range("a65536").get_End(XlDirection.xlUp).Row;
                                        wsheet.get_Range("b1").Value = "清单数量";
                                        wsheet.get_Range("c1").Value = "安装图数量";
                                        for (int i = 2; i < h + 1; i++)
                                        {
                                            string code = (string)wsheet.get_Range("a" + i).Value;
                                            if (dic.ContainsKey(code))
                                            {
                                                wsheet.get_Range("c" + i).Value = dic[code];
                                                if (wsheet.get_Range("c" + i).Value != wsheet.get_Range("b" + i).Value)
                                                {
                                                    wsheet.get_Range("a" + i + ":" + "c" + i).Interior.Color = 255;
                                                }
                                                dic.Remove(code);
                                            }
                                            else
                                            {
                                                wsheet.get_Range("c" + i).Value = 0;
                                                wsheet.get_Range("a" + i + ":" + "c" + i).Interior.Color = 255;
                                            }
                                        }

                                        if (dic.Keys.Count > 0)
                                        {
                                            int j = 2;
                                            wsheet.get_Range("e1").Value = "模板编码（安装图有但清单没有）";
                                            wsheet.get_Range("f1").Value = "数量";
                                            foreach (string item in dic.Keys)
                                            {
                                                wsheet.get_Range("e" + j).Value = item;
                                                wsheet.get_Range("f" + j).Value = dic[item];
                                                wsheet.get_Range("e" + j + ":" + "f" + j).Interior.Color = 255;
                                                j += 1;
                                            }
                                        }
                                        exl.Application.DisplayAlerts = false;
                                        setFormat();

                                        return;

                                    }
                                    else if(g== wbook.Sheets.Count)
                                    {
                                        MessageBox.Show("工作表名称错误，请命名为‘核对数量’！");
                                    }
                                }
                                return;
                            }
                            else if(k == exl.Workbooks.Count)
                            {
                                MessageBox.Show("不存在名称为：‘总量’的工作薄");
                            }
                        }
                   
                    }



                }

            }
            catch (Exception E)
            {
                MyPlugin.ExceptionWrit(E);
            }               
            finally
            {
                exl.Application.DisplayAlerts = true;            
            }

            wbook.Save();           

        }

        private void setFormat()
        {
            wsheet.Sort.SortFields.Clear();
            wsheet.Sort.SortFields.Add(wsheet.get_Range("a2:a" + h), XlSortOn.xlSortOnCellColor,
                XlSortOrder.xlAscending,Type.Missing,XlSortDataOption.xlSortNormal).SortOnValue.Color = 255;
            wsheet.Sort.SetRange(wsheet.get_Range("a1:c" + h));
            wsheet.Sort.Header = XlYesNoGuess.xlYes;
            wsheet.Sort.MatchCase = false;
            wsheet.Sort.Orientation = XlSortOrientation.xlSortColumns;
            wsheet.Sort.SortMethod = XlSortMethod.xlPinYin;
            wsheet.Sort.Apply();
        }

    }

    class EXLf
    {
        private Workbook wbook = null;
        private Worksheet wsheet = null;
        private ExlApp exl;
        private int h;
       // private static int ope = 0;
       // Dictionary<string, Dictionary<string, int>> mab = new Dictionary<string, Dictionary<string, int>>();
        public EXLf()
        {
            try
            {

                int l = Process.GetProcessesByName("EXCEL").Length;
                if (l > 0)
                {
                    exl = (ExlApp)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    //exl.Visible = false;
                }
                else
                {
                    exl = new Microsoft.Office.Interop.Excel.Application();
                    //exl.Visible = false;
                }         

            }
            catch (Exception E)
            {
                MyPlugin.ExceptionWrit(E);

            }

        }

        private void GetEndRow(string c)
        {
            h = wsheet.get_Range(c + 65536).get_End(XlDirection.xlUp).Row + 1;
        }

        public bool evW(Dictionary<string, Dictionary<string, int>> mab1,string[] dirs)
        {
            try
            {
                exl.DisplayAlerts = false;
                exl.ScreenUpdating = false;

                foreach (string dir in dirs)
                {

                    wbook = exl.Workbooks.Open(dir);
                 
                    int cou = 0;

                    foreach (Worksheet item in wbook.Sheets)
                    {
                        wsheet = item;
                        string mbb = wsheet.get_Range("b2").Value;
                      
                        if (mbb == "模板编号" && mab1.ContainsKey(wsheet.Name))
                        {
                            GetEndRow("b");
                            for (int i = 3; i < h; i++)
                            {
                                string mb = wsheet.get_Range("b" + i.ToString()).Value;
                                int mbn = Convert.ToInt32(wsheet.get_Range("c" + i.ToString()).Value);

                                if (mab1[wsheet.Name].ContainsKey(mb))
                                {
                                    cou = mab1[wsheet.Name][mb] - mbn;
                                   // mab1[wsheet.Name].Remove(mb);
                                    if (cou != 0)
                                    {
                                        mab1[wsheet.Name][mb] = cou;
                                      //  keyVa.Add(mb, cou);
                                    }
                                    else
                                    {
                                        mab1[wsheet.Name].Remove(mb);
                                        continue;
                                    }
                                  
                                }
                                else
                                {
                                    cou = - mbn;                                 

                                    if (cou != 0)
                                    {
                                        mab1[wsheet.Name][mb] = cou;
                                        //keyVa.Add(mb, cou);
                                    }
                                }
                            }

                           // if (mab1[wsheet.Name].Count>0)
                          //  {
                          //      foreach (string item2 in mab1[wsheet.Name].Keys)
                           //     {
                           //         keyVa.Add(item2, mab1[wsheet.Name][item2]);
                          //      }
                         //   }
                          //      if (keyVa.Count != 0)
                           // {
                          //      mab.Add(wsheet.Name, keyVa);
                          //  }

                        }
                        else if(mbb == "模板编号" && !mab1.ContainsKey(wsheet.Name))
                        {
                            Dictionary<string, int> keyVa = new Dictionary<string, int>();

                            GetEndRow("b");
                            for (int i = 3; i < h ; i++)
                            {
                                string mb = wsheet.get_Range("b" + i.ToString()).Value;
                                int mbn =Convert.ToInt32("-" + Convert.ToString( Convert.ToInt32( wsheet.get_Range("c" + i.ToString()).Value)));
                                keyVa.Add(mb, mbn);                               
                            }

                            if (keyVa.Count != 0)
                            {
                                mab1.Add(wsheet.Name, keyVa);
                                //mab.Add(wsheet.Name, keyVa);
                            }
                        }
                        else
                        {
                            continue;
                        }

                    }
                   
                    wbook.Close();
                }

                exl.DisplayAlerts = true;
                exl.ScreenUpdating = true;

                Dictionary<string, Dictionary<string, int>> mab1copy = new Dictionary<string, Dictionary<string, int>>(mab1);

                foreach (string item in mab1.Keys)
                {
                    if (mab1[item].Count == 0)
                    {
                        mab1copy.Remove(item);
                    }
                }

                if (mab1copy.Count!=0)
                {
                  
                    SetW(mab1copy);
                    return false;
                }
                return true;
            
            }
            catch (Exception E)
            {
                exl.ScreenUpdating = true;
                exl.DisplayAlerts =true;
                MyPlugin.ExceptionWrit(E);
                return false;

            }
        }

        private void SetW(Dictionary<string, Dictionary<string, int>> mab)
        {
            try
            {
            wbook = exl.Workbooks.Add();
            wsheet = wbook.Sheets["Sheet1"] as Worksheet;
            wsheet.get_Range("a1").Value = "部位";
            wsheet.get_Range("b1").Value = "模板编号";
            wsheet.get_Range("c1").Value = "安装图数量-清单数量";

                int row = 2;

                foreach (string item in mab.Keys)
                {
                    foreach (string item1 in mab[item].Keys)
                    {
                        wsheet.get_Range("a" + row.ToString()).Value = item;
                        wsheet.get_Range("b" + row.ToString()).Value = item1;
                        wsheet.get_Range("c" + row.ToString()).Value = mab[item][item1];
                        row += 1;
                    }
                }
                wsheet.Name = "差异表";
                exl.Visible = true;
                //  wbook.Windows.Application.Visible = true;
                //  wbook.Save();

            }
            catch (Exception E)
            {

                MyPlugin.ExceptionWrit(E);

            }
        }

        public void SetV(Dictionary<string,Dictionary<string,int>> codes)
        {
            try
            {

                GetEndRow("a");

            int row = h;

            foreach (string item in codes.Keys)
            {
                foreach (var item1 in codes[item].Keys)
                {
                    wsheet.get_Range("a1").Value = item;
                    wsheet.get_Range("b1").Value = item1;
                    wsheet.get_Range("c1").Value = codes[item][item1];
                }
               
            }

            wbook.Save();

        }
            catch (Exception E)
            {

                MyPlugin.ExceptionWrit(E);

            }

}
      
    }

    class NPOIExcel
    {
        private IWorkbook workbook = null;

        private string pa = null;

        public NPOIExcel(string pa)
        {
            this.pa = pa;
        }

        public bool Calv(Dictionary<string, Dictionary<string, int>> mab1, string[] dirs)
        {
            try
            {
                foreach (string dir in dirs)
                {
                    string extension = System.IO.Path.GetExtension(dir);
                    FileStream fs = File.OpenRead(dir);

                    if (extension.Equals(".xls"))
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    else
                    {
                        //把xlsx文件中的数据写入wk中
                        workbook = new XSSFWorkbook(fs);
                    }

                    fs.Close();

                    int cou = 0;

                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        ISheet sheet = workbook.GetSheetAt(i);
                        IRow row = sheet.GetRow(1);
                        string mbb = row.GetCell(1).ToString() ;
                        if (mbb == "模板编号" && mab1.ContainsKey(sheet.SheetName))
                        {
                            int rowNu = sheet.LastRowNum;
                            for (int j = 2; j < rowNu; j++)
                            {
                                IRow rowA = sheet.GetRow(j);

                                string mb = rowA.GetCell(1).ToString();
                                int mbn = Convert.ToInt32(rowA.GetCell(2).ToString());

                                if (mab1[sheet.SheetName].ContainsKey(mb))
                                {
                                    cou = mab1[sheet.SheetName][mb] - mbn;

                                    if (cou != 0)
                                    {
                                      mab1[sheet.SheetName][mb] = cou;                                    
                                    }
                                   else
                                    {
                                       mab1[sheet.SheetName].Remove(mb);
                                       continue;
                                    }

                                }
                                else
                                {
                                    cou = -mbn;

                                    if (cou != 0)
                                    {
                                        mab1[sheet.SheetName][mb] = cou;                                      
                                    }
                                }

                            }
                        }
                        else if (mbb == "模板编号" && !mab1.ContainsKey(sheet.SheetName))
                        {
                            Dictionary<string, int> keyVa = new Dictionary<string, int>();
                            int rowNu = sheet.LastRowNum;

                            for (int j = 2; j < rowNu; j++)
                            {
                                IRow rowA = sheet.GetRow(j);
                                string mb = rowA.GetCell(1).ToString();
                                int mbn = Convert.ToInt32("-" + rowA.GetCell(2).ToString());
                                keyVa.Add(mb, mbn);
                            }
                            if (keyVa.Count != 0)
                            {
                                mab1.Add(sheet.SheetName, keyVa);
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }                
                }

                Dictionary<string, Dictionary<string, int>> mab1copy = new Dictionary<string, Dictionary<string, int>>(mab1);

                foreach (string item in mab1.Keys)
                {
                    if (mab1[item].Count == 0)
                    {
                        mab1copy.Remove(item);
                    }
                }
                if (mab1copy.Count != 0)
                {
                    SetW(mab1copy);
                    return false;
                }
                return true;
            }
            catch (Exception E)
            {

                MyPlugin.ExceptionWrit(E);
                return false;

            }
        }

        private void SetW(Dictionary<string, Dictionary<string, int>> mab)
        {
            try
            {
                 HSSFWorkbook  workbook1 = new HSSFWorkbook();

                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();

                workbook1.DocumentSummaryInformation = dsi;
                workbook1.SummaryInformation = si;

                ISheet wsheet = workbook1.CreateSheet("差异表");

                wsheet.SetColumnWidth(2, 5000);

                IRow row = wsheet.CreateRow(0);

                row.Height =300;
                
                row.CreateCell(0).SetCellValue("部位");
                row.CreateCell(1).SetCellValue("模板编号");
                row.CreateCell(2).SetCellValue("安装图数量-清单数量");

                int rowv = 1;

                foreach (string item in mab.Keys)
                {
                    foreach (string item1 in mab[item].Keys)
                    {
                        row = wsheet.CreateRow(rowv);
                        row.Height =300;

                        row.CreateCell(0).SetCellValue(item);
                        row.CreateCell(1).SetCellValue(item1);
                        row.CreateCell(2).SetCellValue(mab[item][item1]);
                        rowv += 1;
                    }
                }

                string posi = pa + "\\差异表.xls";

                FileStream fs = new FileStream(posi, FileMode.Create);

                workbook1.Write(fs);

                fs.Close();
                workbook1.Close();

                try
                {
                     ExlApp exl;
                     int l = Process.GetProcessesByName("EXCEL").Length;
                    if (l > 0)
                    {
                        exl = (ExlApp)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                       
                    }
                    else
                    {
                        exl = new Microsoft.Office.Interop.Excel.Application();
                   
                    }

                    exl.Workbooks.Open(posi);
                    exl.Visible = true;

                }
                catch (Exception E)
                {
                    MyPlugin.ExceptionWrit(E);

                }

                //exl.Visible = true;
                //  wbook.Windows.Application.Visible = true;
                //  wbook.Save();

            }
            catch (Exception E)
            {

                MyPlugin.ExceptionWrit(E);

            }
        }

    }
}
