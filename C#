NPOI读取excel文件

MemoryStream ms = new MemoryStream();    //创建内存流用于写入文件       
IWorkbook workbook = new HSSFWorkbook();   //创建Excel工作部   
ISheet sheet = workbook.CreateSheet("EquipBill");//创建工作表
IRow row = sheet.CreateRow(sheet.LastRowNum);//在工作表中添加一行
ICell cell = row.CreateCell(0);//创建单元格
cell1.SetCellValue("领用单位");//赋值

workbook.Write(ms);//将Excel写入流
ms.Flush();
ms.Position = 0;

FileStream dumpFile = new FileStream(“demo.xls”, FileMode.Create, FileAccess.ReadWrite,FileShare.ReadWrite);
ms.WriteTo(dumpFile);//将流写入文件

NPOI追加excel文件
FileStream fs = new FileStream(“demo.xls”, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);//读取流

POIFSFileSystem ps=new POIFSFileSystem(fs);//需using NPOI.POIFS.FileSystem;
IWorkbook workbook = new HSSFWorkbook(ps);
ISheet sheet = workbook.GetSheetAt(0);//获取工作表
IRow row = sheet.GetRow(0); //得到表头
FileStream fout = new FileStream(“demo.xls”, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);//写入流
row = sheet.CreateRow((sheet.LastRowNum + 1));//在工作表中添加一行

ICell cell1 = row.CreateCell(0);
cell1.SetCellValue(“测试数据”);//赋值

fout.Flush();
workbook.Write(fout);//写入文件
workbook = null;
fout.Close();
